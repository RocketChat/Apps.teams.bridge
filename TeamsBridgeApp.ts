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
import type { IPostRoomUserJoined, IPreRoomUserLeave, IRoom, IRoomUserJoinedContext, IRoomUserLeaveContext } from '@rocket.chat/apps-engine/definition/rooms';
import type { IJobContext } from '@rocket.chat/apps-engine/definition/scheduler';
import { RoomTypeFilter, UIActionButtonContext } from '@rocket.chat/apps-engine/definition/ui';
import type {
	IUIKitResponse,
	UIKitActionButtonInteractionContext,
	UIKitBlockInteractionContext,
	UIKitViewSubmitInteractionContext,
} from '@rocket.chat/apps-engine/definition/uikit';
import type { IFileUploadContext, IPreFileUpload } from '@rocket.chat/apps-engine/definition/uploads';
import type { IPostUserDeleted, IUserContext } from '@rocket.chat/apps-engine/definition/users';

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
import { notifyRocketChatUserInRoomAsync } from './lib/Notifier';
import { OAuthNonce, Room, SubscriptionRenewalJob, WebhookSecret } from './lib/PersistHelper';
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
import { LoginAppUserSlashCommand } from './slashcommands/LoginAppUserSlashCommand';
import { LoginTeamsSlashCommand } from './slashcommands/LoginTeamsSlashCommand';
import { LogoutAppUserSlashCommand } from './slashcommands/LogoutAppUserSlashCommand';
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

		if (data.actionId === UIActionId.AddTeamsUserButtonClicked) {
			const appUser = await read.getUserReader().getAppUser();

			if (!appUser) {
				throw new Error('App user not found');
			}

			const isBridged = await Room.isBridged(read, data.room.id);
			if (!isBridged) {
				await notifyRocketChatUserInRoomAsync(RoomNotBridgedHintMessageText, appUser, data.user, data.room, read.getNotifier());
				return { success: true };
			}

			await openAddTeamsUserContextualBarBlocksAsync(data.triggerId, data.room, data.user, appUser, read, modify, http, persistence, this);
		}

		if (data.actionId === UIActionId.ViewTeamsMembersButtonClicked) {
			const appUser = await read.getUserReader().getAppUser();

			if (!appUser) {
				throw new Error('App user not found');
			}

			const isBridged = await Room.isBridged(read, data.room.id);
			if (!isBridged) {
				await notifyRocketChatUserInRoomAsync(RoomNotBridgedHintMessageText, appUser, data.user, data.room, read.getNotifier());
				return { success: true };
			}

			await openViewTeamsMembersContextualBarAsync(data.triggerId, data.room, data.user, read, modify, http, persistence, this);
		}

		return {
			success: true,
		};
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
			configuration.slashCommands.provideSlashCommand(new LoginAppUserSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new LogoutAppUserSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new ViewTeamsMembersSlashCommand(this)),
			configuration.slashCommands.provideSlashCommand(new BridgeStatusSlashCommand(this)),
		]);

		// Register API endpoints
		await configuration.api.provideApi({
			visibility: ApiVisibility.PUBLIC,
			security: ApiSecurity.UNSECURE,
			endpoints: [new AuthenticationEndpoint(this), new SubscriberEndpoint(this)],
		});

		// Config context menu item - Add Teams user
		configuration.ui.registerButton({
			actionId: UIActionId.AddTeamsUserButtonClicked,
			labelI18n: 'action_button_label_add_teams_user',
			context: UIActionButtonContext.ROOM_ACTION,
			when: {
				roomTypes: [RoomTypeFilter.PRIVATE_DISCUSSION, RoomTypeFilter.PRIVATE_CHANNEL, RoomTypeFilter.PRIVATE_TEAM],
			},
		});

		// Config context menu item - View Teams members
		configuration.ui.registerButton({
			actionId: UIActionId.ViewTeamsMembersButtonClicked,
			labelI18n: 'action_button_label_view_teams_members',
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
