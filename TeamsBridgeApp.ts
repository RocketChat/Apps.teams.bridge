import { App } from '@rocket.chat/apps-engine/definition/App';
import type {
	IAppAccessors,
	IAppInstallationContext,
	IAppUninstallationContext,
	IConfigurationExtend,
	IConfigurationModify,
	IEnvironmentRead,
	IHttp,
	ILogger,
	IMessageBuilder,
	IModify,
	IPersistence,
	IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import { ApiSecurity, ApiVisibility } from '@rocket.chat/apps-engine/definition/api';
import type {
	IMessage,
	IMessageDeleteContext,
	IPostMessageDeleted,
	IPostMessageSent,
	IPostMessageUpdated,
	IPreMessageSentModify,
	IPreMessageSentPrevent,
} from '@rocket.chat/apps-engine/definition/messages';
import type { IAppInfo } from '@rocket.chat/apps-engine/definition/metadata';
import { UserNotAllowedException } from '@rocket.chat/apps-engine/definition/exceptions';
import type { IPostRoomUserJoined, IPreRoomUserJoined, IPreRoomUserLeave, IRoom, IRoomUserJoinedContext, IRoomUserLeaveContext } from '@rocket.chat/apps-engine/definition/rooms';
import type { IJobContext } from '@rocket.chat/apps-engine/definition/scheduler';
import { RoomTypeFilter, UIActionButtonContext } from '@rocket.chat/apps-engine/definition/ui';
import type {
	IUIKitResponse,
	UIKitActionButtonInteractionContext,
	UIKitBlockInteractionContext,
	UIKitViewSubmitInteractionContext,
} from '@rocket.chat/apps-engine/definition/uikit';
import type { IFileUploadContext, IPreFileUpload } from '@rocket.chat/apps-engine/definition/uploads';
import type { IPostUserDeleted, IUser, IUserContext } from '@rocket.chat/apps-engine/definition/users';

import { settings } from './config/Settings';
import { AuthenticationEndpoint } from './endpoints/AuthenticationEndpoint';
import { SubscriberEndpoint } from './endpoints/SubscriberEndpoint';
import {
	IncomingNotificationProcessorId,
	OAuthNonceCleanupInterval,
	OAuthNonceCleanupJobId,
	RegistrationAutoRenewInterval,
	RegistrationAutoRenewSchedulerId,
	SubscriberEndpointPath,
	UIActionId,
	UIElementId,
	WebhookSecretCreationJobId,
	RecentActivityCleanupJobId,
	RecentActivityCleanupInterval,
	RoomNotBridgedHintMessageText,
} from './lib/Const';
import {
	handleAddTeamsUserContextualBarSubmitAsync,
	handlePostMessageDeletedAsync,
	handlePostMessageSentAsync,
	handlePostMessageUpdatedAsync,
	handlePostRoomUserJoinedAsync,
	handlePreFileUploadAsync,
	handlePreMessageSentPreventAsync,
	handlePreRoomUserLeaveAsync,
	handleUninstallApp,
	handleUserRegistrationAutoRenewAsync,
	handlePostUserDeletedAsync,
	handlePreMessageSentModifyAsync,
} from './lib/EventHandler';
import { MapIdentityRcUserBlockId, MapIdentityRcUsernameBlockId, MapIdentityTeamsUserBlockId, openMapIdentityContextualBarAsync } from './lib/MapIdentityInterface';
import { isAppUserLoggedInAsync, performAppUserLoginAsync, performAppUserLogoutAsync } from './lib/AppUserSession';
import { BridgeMapBlockPrefix, buildBridgeMessageView, buildChannelPickerViewAsync, buildChatPickerViewAsync, buildMemberMappingViewAsync, openBridgeRoomContextualBarAsync } from './lib/BridgeRoomInterface';
import { deleteSubscriptionAsync, listMyChatsAsync, listSubscriptionsAsync, listTeamChannelsAsync, subscribeToChannelMessagesAsync } from './lib/MicrosoftGraphApi';
import { getUserAccessTokenAsync } from './lib/AuthHelper';
import { buildHubViewAsync, buildUnlinkConfirmView, openHubViewAsync } from './lib/HubInterface';
import { restoreMappingBackupAsync, syncMappingBackupAsync } from './lib/MappingBackup';
import { notifyRocketChatUserInRoomAsync } from './lib/Notifier';
import { OAuthNonce, Room, SubscriptionRenewalJob, UserMapping, WebhookSecret } from './lib/PersistHelper';
import { getRocketChatAppEndpointUrl } from './lib/UrlHelper';
import {
	decodeButtonState,
	getRoomIdFromActionId,
	isActionId,
	openAddTeamsUserContextualBarBlocksAsync,
	openViewTeamsMembersContextualBarAsync,
	updateAddTeamsUserContextualBarAsync,
	updateViewTeamsMembersContextualBarAsync,
} from './lib/UserInterfaceHelper';
import { handleWebhookSecretCreationAsync } from './lib/handlers/handleWebhookSecretCreation';
import { handleInboundNotificationAsync } from './lib/inboundNotification/handleInboundNotificationAsync';
import { RecentActivity } from './lib/persistence';
import { AddUserSlashCommand } from './slashcommands/AddUserSlashCommand';
import { BridgeStatusSlashCommand } from './slashcommands/BridgeStatusSlashCommand';
import { LoginTeamsSlashCommand } from './slashcommands/LoginTeamsSlashCommand';
import { LogoutTeamsSlashCommand } from './slashcommands/LogoutTeamsSlashCommand';
import { ResubscribeMessages } from './slashcommands/ResubscriptionMessages';
import { SetupVerificationSlashCommand } from './slashcommands/SetupVerificationSlashCommand';
import { ViewTeamsMembersSlashCommand } from './slashcommands/ViewTeamsMembersSlashCommand';

export class TeamsBridgeApp
	extends App
	implements
		IPreMessageSentPrevent,
		IPostMessageSent,
		IPostMessageUpdated,
		IPostMessageDeleted,
		IPreFileUpload,
		IPreMessageSentModify,
		IPreRoomUserLeave,
		IPreRoomUserJoined,
		IPostRoomUserJoined,
		IPostUserDeleted
{
	constructor(info: IAppInfo, logger: ILogger, accessors: IAppAccessors) {
		super(info, logger, accessors);
	}

	protected recentActivityCleanupJob = async (_jobContext: IJobContext, read: IRead, _modify: IModify, _http: IHttp, persis: IPersistence): Promise<void> => {
		const deletedCount = await RecentActivity.deleteStale(read, persis);
		this.getLogger().info(`Deleted ${deletedCount} stale recent activities.`);
	};

	async getSettingValueById(id: string) {
		return this.getAccessors().environmentReader.getSettings().getValueById(id);
	}

	async onInstall(_context: IAppInstallationContext, _read: IRead, _http: IHttp, persistence: IPersistence, _modify: IModify): Promise<void> {
		await WebhookSecret.create({ persistence });
	}

	public async executePreMessageSentModify(
		message: IMessage,
		builder: IMessageBuilder,
		read: IRead,
		http: IHttp,
		persistence: IPersistence,
	): Promise<IMessage> {
		return handlePreMessageSentModifyAsync({
			app: this,
			message,
			builder,
			read,
			persistence,
			http,
		});
	}

	async onEnable(_environment: IEnvironmentRead, configurationModify: IConfigurationModify): Promise<boolean> {
		try {
			await configurationModify.scheduler.scheduleOnce({
				id: WebhookSecretCreationJobId,
				when: new Date(),
				data: { from: 'ScheduleOnce/Immediate' },
			});

			await configurationModify.scheduler.scheduleOnce({
				id: RegistrationAutoRenewSchedulerId,
				when: new Date(Date.now() + 5000),
				data: { from: 'ScheduleOnce/5seconds' },
			});

			await configurationModify.scheduler.scheduleRecurring({
				id: RegistrationAutoRenewSchedulerId,
				interval: RegistrationAutoRenewInterval,
				data: { from: 'ScheduleRecurring' },
				skipImmediate: true,
			});

			await configurationModify.scheduler.scheduleRecurring({
				id: OAuthNonceCleanupJobId,
				interval: OAuthNonceCleanupInterval,
				skipImmediate: true,
			});
			await configurationModify.scheduler.scheduleRecurring({
				id: RecentActivityCleanupJobId,
				interval: RecentActivityCleanupInterval,
				skipImmediate: true,
			});
		} catch (e) {
			this.getLogger().error(e);
		}
		return true;
	}

	async onDisable(configurationModify: IConfigurationModify): Promise<void> {
		await configurationModify.scheduler.cancelAllJobs();
	}

	public async onUninstall(_context: IAppUninstallationContext, read: IRead, http: IHttp, persistence: IPersistence, modify: IModify): Promise<void> {
		return handleUninstallApp({
			read,
			http,
			modify,
			app: this,
			persistence,
		});
	}

	public async executePreMessageSentPrevent(message: IMessage, read: IRead, http: IHttp, persistence: IPersistence): Promise<boolean> {
		return handlePreMessageSentPreventAsync({
			app: this,
			message,
			read,
			persistence,
			http,
		});
	}

	public async executePostMessageSent(message: IMessage, read: IRead, http: IHttp, persistence: IPersistence, modify: IModify): Promise<void> {
		await handlePostMessageSentAsync({
			app: this,
			message,
			read,
			persistence,
			http,
			modify,
		});
	}

	public async executePostMessageUpdated(message: IMessage, read: IRead, http: IHttp, persistence: IPersistence, _modify: IModify): Promise<void> {
		await handlePostMessageUpdatedAsync({
			app: this,
			message,
			read,
			persistence,
			http,
		});
	}

	public async executePostMessageDeleted(
		message: IMessage,
		read: IRead,
		http: IHttp,
		persistence: IPersistence,
		_modify: IModify,
		_context: IMessageDeleteContext,
	): Promise<void> {
		await handlePostMessageDeletedAsync({
			app: this,
			message,
			read,
			persistence,
			http,
		});
	}

	public async executePreFileUpload(context: IFileUploadContext, read: IRead, http: IHttp, persistence: IPersistence, _modify: IModify): Promise<void> {
		await handlePreFileUploadAsync({
			app: this,
			context,
			read,
			persistence,
			http,
		});
	}

	public async executePreRoomUserLeave(context: IRoomUserLeaveContext, read: IRead, http: IHttp, persistence: IPersistence): Promise<void> {
		await handlePreRoomUserLeaveAsync({
			app: this,
			context,
			read,
			persistence,
			http,
		});
	}

	// Channel-linked rooms mirror the Teams channel roster: members must be added on the
	// Teams side (then mapped). Block RC-side adds of anyone who is not the bot or an
	// already-mapped Teams member.
	public async executePreRoomUserJoined(context: IRoomUserJoinedContext, read: IRead): Promise<void> {
		const roomRecord = await Room.findByRCRoomId(read, context.room.id);
		if (!roomRecord?.teamsTeamId) {
			return;
		}

		const appUser = await read.getUserReader().getAppUser();
		if (appUser && context.joiningUser.id === appUser.id) {
			return;
		}

		const mapping = await UserMapping.findByRCUserId(read, context.joiningUser.id);
		if (mapping) {
			return;
		}

		throw new UserNotAllowedException(
			'This room is linked to a Microsoft Teams channel. Add the person to the channel on the Teams side; once they are mapped via Teams Bridge they can join here.',
		);
	}

	public async executePostRoomUserJoined(
		context: IRoomUserJoinedContext,
		read: IRead,
		http: IHttp,
		persistence: IPersistence,
		modify: IModify,
	): Promise<void> {
		await handlePostRoomUserJoinedAsync({
			app: this,
			context,
			read,
			persistence,
			http,
			modify,
		});
	}

	// eslint-disable-next-line @typescript-eslint/no-empty-function
	public async executePostUserCreated(_context: IUserContext, _read: IRead, _http: IHttp, _persistence: IPersistence, _modify: IModify): Promise<void> {}

	public async executePostUserDeleted(context: IUserContext, read: IRead, http: IHttp, persistence: IPersistence, modify: IModify): Promise<void> {
		await handlePostUserDeletedAsync({
			app: this,
			context,
			read,
			persistence,
			http,
			modify,
		});
	}

	public async executeActionButtonHandler(
		context: UIKitActionButtonInteractionContext,
		read: IRead,
		http: IHttp,
		persistence: IPersistence,
		modify: IModify,
	): Promise<IUIKitResponse> {
		const data = context.getInteractionData();

		if (data.actionId === UIActionId.TeamsBridgeHubButtonClicked) {
			await openHubViewAsync(data.triggerId, data.room, data.user, read, http, persistence, modify, this);
		}

		return {
			success: true,
		};
	}

	// Shared router used by the hub's in-modal buttons. Login/logout and bridging need no
	// bridged-room guard; the member tools do.
	private async routeTeamsBridgeToolAsync(
		actionId: string,
		triggerId: string,
		room: IRoom,
		operator: IUser,
		read: IRead,
		modify: IModify,
		http: IHttp,
		persistence: IPersistence,
	): Promise<void> {
		const appUser = await read.getUserReader().getAppUser();
		if (!appUser || !room) {
			return;
		}

		if (isActionId(actionId, UIActionId.AppUserSessionButtonClicked)) {
			const sessionOptions = { read, modify, http, persistence, app: this, operator, room };
			const loggedIn = await isAppUserLoggedInAsync({ read, http, persistence, app: this });
			if (loggedIn) {
				await performAppUserLogoutAsync(sessionOptions);
			} else {
				await performAppUserLoginAsync(sessionOptions);
			}
			return;
		}

		if (isActionId(actionId, UIActionId.BridgeRoomButtonClicked)) {
			await openBridgeRoomContextualBarAsync(triggerId, room, operator, read, modify, http, persistence, this);
			return;
		}

		const isBridged = await Room.isBridged(read, room.id);
		if (!isBridged) {
			await notifyRocketChatUserInRoomAsync(RoomNotBridgedHintMessageText, appUser, operator, room, read.getNotifier());
			return;
		}

		if (isActionId(actionId, UIActionId.AddTeamsUserButtonClicked)) {
			await openAddTeamsUserContextualBarBlocksAsync(triggerId, room, operator, appUser, read, modify, http, persistence, this);
		} else if (isActionId(actionId, UIActionId.ViewTeamsMembersButtonClicked)) {
			await openViewTeamsMembersContextualBarAsync(triggerId, room, operator, read, modify, http, persistence, this);
		} else if (isActionId(actionId, UIActionId.MapIdentityButtonClicked)) {
			await openMapIdentityContextualBarAsync(triggerId, room, operator, read, modify, http, persistence, this);
		}
	}

	public async executeBlockActionHandler(
		context: UIKitBlockInteractionContext,
		read: IRead,
		http: IHttp,
		persistence: IPersistence,
		modify: IModify,
	): Promise<IUIKitResponse> {
		const { actionId, value, room } = context.getInteractionData();
		const roomId = room?.id ?? getRoomIdFromActionId(actionId) ?? '';

		if (isActionId(actionId, UIActionId.TeamsUserSearchInput)) {
			// Fires on every keystroke (ON_CHARACTER_ENTERED dispatch).
			// value = current text in the search input; room = the open room.
			const updatedView = await updateAddTeamsUserContextualBarAsync({
				actionId,
				value,
				roomId,
				read,
				http,
				persistence,
				app: this,
				modify,
			});

			if (updatedView) {
				return context.getInteractionResponder().updateContextualBarViewResponse(updatedView);
			}
		}

		if (isActionId(actionId, UIActionId.TeamsUserLoadMore) && value) {
			// value = base64 encoded { nextLink, loadedUsers, roomId }.
			const { roomId } = decodeButtonState(value);
			const updatedView = await updateAddTeamsUserContextualBarAsync({
				actionId,
				value,
				roomId,
				read,
				http,
				persistence,
				app: this,
				modify,
			});
			if (updatedView) {
				return context.getInteractionResponder().updateContextualBarViewResponse(updatedView);
			}
		}

		if (isActionId(actionId, UIActionId.ViewMembersLoadMore) && value) {
			// value = base64 encoded { nextLink, loadedMembers, threadId, total }.
			const updatedView = await updateViewTeamsMembersContextualBarAsync({
				value,
				read,
				http,
				persistence,
				app: this,
				modify,
				roomId,
			});
			if (updatedView) {
				return context.getInteractionResponder().updateContextualBarViewResponse(updatedView);
			}
		}

		const interaction = context.getInteractionData();
		const actor = interaction.user;
		const triggerId = interaction.triggerId;
		const actionRoom = interaction.room ?? (roomId ? await read.getRoomReader().getById(roomId) : undefined);

		// Hub: Refresh re-renders the panel in place (no page reload needed)
		if (actionRoom && isActionId(actionId, UIActionId.TeamsBridgeHubRefresh)) {
			const hub = await buildHubViewAsync({ read, http, persistence, modify, app: this, room: actionRoom });
			return context.getInteractionResponder().updateContextualBarViewResponse(hub);
		}

		// Hub: restore identity mappings from the backup setting
		if (actionRoom && actor && isActionId(actionId, UIActionId.TeamsBridgeRestoreMappings)) {
			const appUser = await read.getUserReader().getAppUser();
			const result = await restoreMappingBackupAsync({ read, persistence, http, app: this });
			if (appUser) {
				const text = result.error
					? `⚠️ Restore failed: ${result.error}`
					: `✅ Restored ${result.restored} mapping(s)${result.skipped ? `, skipped ${result.skipped} (user not found)` : ''} from the backup setting.`;
				await notifyRocketChatUserInRoomAsync(text, appUser, actor, actionRoom, read.getNotifier());
			}
			const hub = await buildHubViewAsync({ read, http, persistence, modify, app: this, room: actionRoom });
			return context.getInteractionResponder().updateContextualBarViewResponse(hub);
		}

		// Hub: Break Teams map — ask for confirmation first (destructive).
		if (actionRoom && isActionId(actionId, UIActionId.BridgeUnlinkClicked)) {
			const record = await Room.findByRCRoomId(read, actionRoom.id);
			const name = record?.teamsThreadName ?? '';
			return context.getInteractionResponder().updateContextualBarViewResponse(buildUnlinkConfirmView(modify, actionRoom, name));
		}

		// Hub: unlink cancelled — back to the hub.
		if (actionRoom && isActionId(actionId, UIActionId.BridgeUnlinkCancel)) {
			const hub = await buildHubViewAsync({ read, http, persistence, modify, app: this, room: actionRoom });
			return context.getInteractionResponder().updateContextualBarViewResponse(hub);
		}

		// Hub: unlink confirmed — remove the Teams link (and any channel subscription), then re-render.
		if (actionRoom && actor && isActionId(actionId, UIActionId.BridgeUnlinkConfirm)) {
			const record = await Room.findByRCRoomId(read, actionRoom.id);
			const appUser = await read.getUserReader().getAppUser();

			if (record?.teamsTeamId && record.teamsThreadId && appUser) {
				try {
					const token = await getUserAccessTokenAsync({ read, persistence, http, app: this, rocketChatUserId: appUser.id });
					if (token) {
						const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(this.getAccessors(), SubscriberEndpointPath);
						const channelNotificationUrl = `${subscriberEndpointUrl}?userId=${appUser.id}&channelId=${encodeURIComponent(record.teamsThreadId)}&hasClientState=1`;
						const subs = (await listSubscriptionsAsync(http, token, channelNotificationUrl)) ?? [];
						await Promise.all(subs.map((sub) => deleteSubscriptionAsync(http, sub.id, token)));
					}
				} catch (error) {
					this.getLogger().warn(`Failed to delete channel subscription during unlink: ${error}`);
				}
			}

			await Room.unlink(read, persistence, actionRoom.id);

			if (appUser) {
				await notifyRocketChatUserInRoomAsync(
					`🔓 Teams map removed${record?.teamsThreadName ? ` (was *${record.teamsThreadName}*)` : ''}. Messages no longer relay. Use *Map to Teams* to re-link.`,
					appUser,
					actor,
					actionRoom,
					read.getNotifier(),
				);
			}

			const hub = await buildHubViewAsync({ read, http, persistence, modify, app: this, room: actionRoom });
			return context.getInteractionResponder().updateContextualBarViewResponse(hub);
		}

		// Hub: login/logout runs, then the panel re-renders so status updates immediately
		if (actionRoom && actor && isActionId(actionId, UIActionId.AppUserSessionButtonClicked)) {
			await this.routeTeamsBridgeToolAsync(actionId, triggerId, actionRoom, actor, read, modify, http, persistence);
			const hub = await buildHubViewAsync({ read, http, persistence, modify, app: this, room: actionRoom });
			return context.getInteractionResponder().updateContextualBarViewResponse(hub);
		}

		// Hub buttons that open their own surface (bridge / add / view / map)
		if (
			actionRoom &&
			triggerId &&
			actor &&
			(isActionId(actionId, UIActionId.BridgeRoomButtonClicked) ||
				isActionId(actionId, UIActionId.AddTeamsUserButtonClicked) ||
				isActionId(actionId, UIActionId.ViewTeamsMembersButtonClicked) ||
				isActionId(actionId, UIActionId.MapIdentityButtonClicked))
		) {
			await this.routeTeamsBridgeToolAsync(actionId, triggerId, actionRoom, actor, read, modify, http, persistence);
			return context.getInteractionResponder().successResponse();
		}

		// Bridge chooser: auto-create
		if (isActionId(actionId, UIActionId.BridgeAutoCreateClicked)) {
			const rid = getRoomIdFromActionId(actionId);
			const r = rid ? await read.getRoomReader().getById(rid) : undefined;
			const appUser = await read.getUserReader().getAppUser();
			if (r && appUser && actor) {
				const members = await read.getRoomReader().getMembers(r.id);
				if (!members.some((m) => m.id === appUser.id)) {
					const rb = await modify.getUpdater().room(r.id, actor);
					rb.addMemberToBeAddedByUsername(appUser.username);
					await modify.getUpdater().finish(rb);
				}
			}
			return context.getInteractionResponder().updateContextualBarViewResponse(
				buildBridgeMessageView(modify, [
					':white_check_mark: *Bridged (auto-create).*',
					'A Teams chat will be created automatically when a mapped user first posts in this room.',
				]),
			);
		}

		// Bridge chooser: link existing -> chat picker
		if (isActionId(actionId, UIActionId.BridgeLinkExistingClicked)) {
			const rid = getRoomIdFromActionId(actionId);
			const r = rid ? await read.getRoomReader().getById(rid) : undefined;
			if (r) {
				const view = await buildChatPickerViewAsync({ read, http, persistence, modify, app: this, room: r });
				return context.getInteractionResponder().updateContextualBarViewResponse(view);
			}
		}

		// Bridge: a Teams chat was selected -> force-map all its members
		if (isActionId(actionId, UIActionId.BridgeChatSelected) && value) {
			const rid = getRoomIdFromActionId(actionId);
			const r = rid ? await read.getRoomReader().getById(rid) : undefined;
			if (r) {
				const view = await buildMemberMappingViewAsync({ read, http, persistence, modify, app: this, room: r, threadId: value });
				return context.getInteractionResponder().updateContextualBarViewResponse(view);
			}
		}

		// Bridge chooser: link existing channel -> channel picker
		if (isActionId(actionId, UIActionId.BridgeLinkChannelClicked)) {
			const rid = getRoomIdFromActionId(actionId);
			const r = rid ? await read.getRoomReader().getById(rid) : undefined;
			if (r) {
				const view = await buildChannelPickerViewAsync({ read, http, persistence, modify, app: this, room: r });
				return context.getInteractionResponder().updateContextualBarViewResponse(view);
			}
		}

		// Bridge: a Team channel was selected (value = "teamId|channelId") -> force-map all its members
		if (isActionId(actionId, UIActionId.BridgeChannelSelected) && value) {
			const rid = getRoomIdFromActionId(actionId);
			const r = rid ? await read.getRoomReader().getById(rid) : undefined;
			const [teamId, channelId] = value.split('|');
			if (r && teamId && channelId) {
				const view = await buildMemberMappingViewAsync({ read, http, persistence, modify, app: this, room: r, threadId: channelId, teamId });
				return context.getInteractionResponder().updateContextualBarViewResponse(view);
			}
		}

		return context.getInteractionResponder().successResponse();
	}

	public async executeViewClosedHandler(): Promise<IUIKitResponse> {
		return Promise.resolve({
			success: true,
		});
	}

	public async executeViewSubmitHandler(
		context: UIKitViewSubmitInteractionContext,
		read: IRead,
		http: IHttp,
		persistence: IPersistence,
		modify: IModify,
	): Promise<IUIKitResponse> {
		const { user, view } = context.getInteractionData();

		if (view.id === UIElementId.ContextualBarId) {
			const submitActionId = view.submit?.actionId;

			let currentRoom: IRoom | undefined;
			const roomIdFromActionId = submitActionId && getRoomIdFromActionId(submitActionId);
			if (roomIdFromActionId) {
				const room = await read.getRoomReader().getById(roomIdFromActionId);
				if (room) {
					currentRoom = room;
				}
			}

			let teamsUserIdsToSave: string[] | undefined;

			if (view.state) {
				Object.values(view.state).forEach((item) => {
					Object.entries(item).forEach(([key, value]) => {
						if (isActionId(key, UIActionId.TeamsUserNameSearch)) {
							teamsUserIdsToSave = value as string[] | undefined;
						}
					});
				});
			}

			if (teamsUserIdsToSave && currentRoom) {
				await handleAddTeamsUserContextualBarSubmitAsync({
					operator: user,
					room: currentRoom,
					teamsUserIdsToSave,
					read,
					persistence,
					http,
					modify,
					app: this,
				});
			}
		}

		if (view.id === UIElementId.MapIdentityContextualBarId) {
			const state = (view.state ?? {}) as Record<string, Record<string, string>>;
			const selectedRcUserId = state[MapIdentityRcUserBlockId]?.[UIActionId.MapIdentityRcUserSelect];
			const typedUsername = (state[MapIdentityRcUsernameBlockId]?.[UIActionId.MapIdentityRcUsernameInput] ?? '').trim().replace(/^@/, '');
			const teamsUserId = state[MapIdentityTeamsUserBlockId]?.[UIActionId.MapIdentityTeamsUserSelect];

			const roomId = getRoomIdFromActionId(view.submit?.actionId ?? '');
			const room = roomId ? await read.getRoomReader().getById(roomId) : undefined;
			const appUser = await read.getUserReader().getAppUser();

			// A typed username (any RC user) wins over the room-member dropdown.
			let rcUserId = selectedRcUserId;
			if (typedUsername) {
				const typedUser = await read.getUserReader().getByUsername(typedUsername);
				if (!typedUser) {
					if (room && appUser) {
						await notifyRocketChatUserInRoomAsync(
							`No Rocket.Chat user named @${typedUsername} was found. Mapping not saved.`,
							appUser,
							user,
							room,
							read.getNotifier(),
						);
					}
					return { success: true };
				}
				rcUserId = typedUser.id;
			}

			if (rcUserId && teamsUserId) {
				await UserMapping.persist(persistence, rcUserId, teamsUserId);
				await syncMappingBackupAsync({ read, app: this });

				const mappedUser = await read.getUserReader().getById(rcUserId);
				if (room && appUser && mappedUser) {
					// Add the mapped user to the room so inbound attribution works right away
					// (mapped users are allowed through the channel-room join guard).
					const members = await read.getRoomReader().getMembers(room.id);
					if (!members.some((m) => m.id === mappedUser.id)) {
						try {
							const rb = await modify.getUpdater().room(room.id, user);
							rb.addMemberToBeAddedByUsername(mappedUser.username);
							await modify.getUpdater().finish(rb);
						} catch (error) {
							this.getLogger().warn(`Could not add mapped user @${mappedUser.username} to room: ${error}`);
						}
					}
					await notifyRocketChatUserInRoomAsync(
						`Mapped Teams user to Rocket.Chat user @${mappedUser.username} and added them to this room. Their Teams messages will now appear as this user.`,
						appUser,
						user,
						room,
						read.getNotifier(),
					);
				}
			} else if (room && appUser) {
				await notifyRocketChatUserInRoomAsync(
					'Mapping not saved: pick a Teams user and either select a room member or type a Rocket.Chat username.',
					appUser,
					user,
					room,
					read.getNotifier(),
				);
			}
		}

		if (view.id === UIElementId.BridgeRoomContextualBarId) {
			const submitActionId = view.submit?.actionId ?? '';
			const parts = submitActionId.split('--');
			// Channel links carry "teamId|channelId" in the middle segment; chat links carry the thread id.
			const middle = parts[1] ?? '';
			const isChannelLink = middle.includes('|');
			const [linkTeamId, linkChannelId] = isChannelLink ? middle.split('|') : [undefined, undefined];
			const threadId = isChannelLink ? linkChannelId : middle;
			const roomId = getRoomIdFromActionId(submitActionId);
			const state = (view.state ?? {}) as Record<string, Record<string, string>>;

			const mappings: Array<{ teamsUserId: string; rcUserId: string }> = [];
			let totalBlocks = 0;
			for (const [blockId, inner] of Object.entries(state)) {
				if (blockId.startsWith(`${BridgeMapBlockPrefix}--`)) {
					totalBlocks++;
					const teamsUserId = blockId.slice(`${BridgeMapBlockPrefix}--`.length);
					const rcUserId = Object.values(inner ?? {})[0];
					if (teamsUserId && rcUserId) {
						mappings.push({ teamsUserId, rcUserId });
					}
				}
			}

			const appUser = await read.getUserReader().getAppUser();
			const room = roomId ? await read.getRoomReader().getById(roomId) : undefined;

			if (!threadId || !roomId || totalBlocks === 0 || mappings.length < totalBlocks) {
				if (room && appUser) {
					await notifyRocketChatUserInRoomAsync(
						'Link not saved: every Teams member must be mapped to a Rocket.Chat user.',
						appUser,
						user,
						room,
						read.getNotifier(),
					);
				}
				return { success: true };
			}

			for (const m of mappings) {
				await UserMapping.persist(persistence, m.rcUserId, m.teamsUserId);
			}
			await syncMappingBackupAsync({ read, app: this });

			// Resolve a friendly name for the linked target so the hub can display it later.
			let linkedName = '';
			const appUserToken = appUser
				? await getUserAccessTokenAsync({ read, persistence, http, app: this, rocketChatUserId: appUser.id })
				: undefined;
			if (appUserToken) {
				if (isChannelLink && linkTeamId) {
					const channels = (await listTeamChannelsAsync(http, linkTeamId, appUserToken)) ?? [];
					linkedName = channels.find((c) => c.id === threadId)?.displayName ?? '';
				} else {
					const chats = await listMyChatsAsync(http, appUserToken);
					linkedName = chats?.chats.find((c) => c.id === threadId)?.topic ?? '';
				}
			}

			await Room.persist(read, persistence, roomId, threadId, {
				...(isChannelLink && linkTeamId ? { teamsTeamId: linkTeamId } : {}),
				...(linkedName ? { teamsThreadName: linkedName } : {}),
			});
			await Room.setBridgeActive(persistence, read, roomId, true);

			// Channel links need their own Graph change-subscription (the per-user chat
			// subscription does not cover Team channels).
			if (isChannelLink && linkTeamId && threadId && appUser && appUserToken) {
				try {
					const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(this.getAccessors(), SubscriberEndpointPath);
					await subscribeToChannelMessagesAsync({
						http,
						read,
						persis: persistence,
						rocketChatUserId: appUser.id,
						teamId: linkTeamId,
						channelId: threadId,
						subscriberEndpointUrl,
						userAccessToken: appUserToken,
						renewIfExists: true,
					});
				} catch (error) {
					this.getLogger().error(`Failed to subscribe to channel messages: ${error}`);
				}
			}

			if (room && appUser) {
				const members = await read.getRoomReader().getMembers(room.id);
				if (!members.some((mem) => mem.id === appUser.id)) {
					const rb = await modify.getUpdater().room(room.id, user);
					rb.addMemberToBeAddedByUsername(appUser.username);
					await modify.getUpdater().finish(rb);
				}
				await notifyRocketChatUserInRoomAsync(
					`:link: Linked this room to the Teams ${isChannelLink ? 'channel' : 'chat'}${linkedName ? ` *${linkedName}*` : ''} and mapped ${mappings.length} member(s). Messages now relay both ways.${isChannelLink ? ' Members can only be added on the Teams side.' : ''}`,
					appUser,
					user,
					room,
					read.getNotifier(),
				);
			}
		}

		return {
			success: true,
		};
	}

	protected incomingNotificationJob = async (
		jobContext: IJobContext,
		read: IRead,
		modify: IModify,
		http: IHttp,
		persistence: IPersistence,
	): Promise<void> => {
		await handleInboundNotificationAsync({
			app: this,
			read,
			modify,
			http,
			inBoundNotification: jobContext.inBoundNotification,
			persistence,
		});
	};

	protected webhookSecretCreationJob = async (
		_jobContext: IJobContext,
		read: IRead,
		_modify: IModify,
		http: IHttp,
		persistence: IPersistence,
	): Promise<void> => {
		await handleWebhookSecretCreationAsync({
			app: this,
			read,
			http,
			persistence,
		});
	};

	protected registrationRenewalsJob = async (jobContext: IJobContext, read: IRead, _modify: IModify, http: IHttp, persistence: IPersistence) => {
		try {
			this.getLogger().info(`[Teams Bridge] Start renew registrations! (from: ${jobContext.from})`);
			const jobState = await SubscriptionRenewalJob.find({
				persistenceRead: read.getPersistenceReader(),
			});

			if (jobState && jobState.lastStartedJobTimestamp && Date.now() - new Date(jobState.lastStartedJobTimestamp).getTime() < 5 * 60 * 1000) {
				// Job ran less than 5 minutes ago
				this.getLogger().info(`[Teams Bridge] ${RegistrationAutoRenewSchedulerId} Job already ran less than 5 minutes ago. Skipping this run.`);
				return;
			}

			await SubscriptionRenewalJob.persist({
				persistence,
				lastStartedJobTimestamp: new Date(),
			});

			const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(this.getAccessors(), SubscriberEndpointPath);

			await handleUserRegistrationAutoRenewAsync({
				subscriberEndpointUrl,
				read,
				http,
				persistence,
				app: this,
			});
			this.getLogger().info('[Teams Bridge] Finish renew registrations!');
		} catch (error) {
			throw new Error(`[Teams Bridge] Auto renew registration failed with error: ${error}`);
		}
	};

	protected oauthNonceCleanupJob = async (_jobContext: IJobContext, read: IRead, _modify: IModify, _http: IHttp, persistence: IPersistence) => {
		try {
			await OAuthNonce.deleteStale(read, persistence);
		} catch (error) {
			throw new Error(`[Teams Bridge] OAuth nonce cleanup failed with error: ${error}`);
		}
	};

	protected async extendConfiguration(configuration: IConfigurationExtend): Promise<void> {
		// Register app settings
		await Promise.all(settings.map((setting) => configuration.settings.provideSetting(setting)));

		await Promise.all([
			configuration.slashCommands.provideSlashCommand(new SetupVerificationSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new LoginTeamsSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new LogoutTeamsSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new AddUserSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new ResubscribeMessages(this)),
			configuration.slashCommands.provideSlashCommand(new ViewTeamsMembersSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new BridgeStatusSlashCommand(this)),
		]);

		// Register API endpoints
		await configuration.api.provideApi({
			visibility: ApiVisibility.PUBLIC,
			security: ApiSecurity.UNSECURE,
			endpoints: [new AuthenticationEndpoint(this), new SubscriberEndpoint(this)],
		});

		// Single room-action entry point: "Teams Bridge" opens the hub modal; every capability
		// (login/logout, bridge, add/view/map members) is a button inside it.
		configuration.ui.registerButton({
			actionId: UIActionId.TeamsBridgeHubButtonClicked,
			labelI18n: 'action_button_label_teams_bridge_hub',
			context: UIActionButtonContext.ROOM_ACTION,
			when: {
				roomTypes: [RoomTypeFilter.PRIVATE_DISCUSSION, RoomTypeFilter.PRIVATE_CHANNEL, RoomTypeFilter.PRIVATE_TEAM],
			},
		});

		// Config a scheduler for UserAccessToken & Subscription auto renew and start it
		await configuration.scheduler.registerProcessors([
			{
				id: RegistrationAutoRenewSchedulerId,
				processor: this.registrationRenewalsJob,
			},
			{
				id: WebhookSecretCreationJobId,
				processor: this.webhookSecretCreationJob,
			},
			{
				id: IncomingNotificationProcessorId,
				processor: this.incomingNotificationJob,
			},
			{
				id: OAuthNonceCleanupJobId,
				processor: this.oauthNonceCleanupJob,
			},
			{
				id: RecentActivityCleanupJobId,
				processor: this.recentActivityCleanupJob,
			},
		]);
	}
}
