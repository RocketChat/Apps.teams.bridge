import type { IHttp, IModify, IPersistence, IRead, IUIKitSurfaceViewParam } from '@rocket.chat/apps-engine/definition/accessors';
import type { IRoom } from '@rocket.chat/apps-engine/definition/rooms';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';
import { UIKitSurfaceType } from '@rocket.chat/apps-engine/definition/uikit';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { getUserAccessTokenAsync } from './AuthHelper';
import { UIActionId, UIElementId } from './Const';
import { getChannelMembersAsync, getTeamsChatMembersAsync, listJoinedTeamsAsync, listMyChatsAsync, listTeamChannelsAsync } from './MicrosoftGraphApi';
import { notifyRocketChatUserInRoomAsync } from './Notifier';
import { Room } from './PersistHelper';

// blockId = `${BridgeMapBlockPrefix}--${teamsUserId}`; the selected value is the RC user id.
export const BridgeMapBlockPrefix = 'TeamsBridge.BridgeMap';
export const BridgeMapSelectPrefix = 'TeamsBridge.BridgeMapSelect';

const TITLE = 'Bridge room to Teams';

const truncate = (s: string, n = 72): string => (s.length > n ? `${s.slice(0, n - 1)}…` : s);

// A simple section-only view (no submit) used for prompts / errors while the bar is open.
export const buildBridgeMessageView = (modify: IModify, lines: string[]): IUIKitSurfaceViewParam => {
	const blocks = modify.getCreator().getBlockBuilder();
	lines.forEach((line) => blocks.addSectionBlock({ text: blocks.newMarkdownTextObject(line) }));
	return {
		id: UIElementId.BridgeRoomContextualBarId,
		title: blocks.newPlainTextObject(TITLE),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		blocks: blocks.getBlocks(),
	};
};

// Step 1: choose how to bridge (auto-create a new Teams chat, or link an existing one).
export const buildBridgeChooserView = (modify: IModify, room: IRoom): IUIKitSurfaceViewParam => {
	const blocks = modify.getCreator().getBlockBuilder();
	blocks.addSectionBlock({ text: blocks.newMarkdownTextObject('*How should this room bridge to Microsoft Teams?*') });
	blocks.addSectionBlock({
		text: blocks.newMarkdownTextObject(
			'*Auto-create* — the bot creates a new Teams chat automatically when a mapped user first posts.\n' +
				'*Link existing chat* — connect this room to a Teams group chat the bot account is already in.\n' +
				'*Link existing channel* — connect this room to a channel of a Team the bot account belongs to.\n' +
				'Linking requires mapping every Teams member to a Rocket.Chat user.',
		),
	});
	blocks.addActionsBlock({
		blockId: 'TeamsBridge.BridgeChooser',
		elements: [
			blocks.newButtonElement({
				actionId: `${UIActionId.BridgeAutoCreateClicked}--${room.id}`,
				text: blocks.newPlainTextObject('Auto-create Teams chat'),
			}),
			blocks.newButtonElement({
				actionId: `${UIActionId.BridgeLinkExistingClicked}--${room.id}`,
				text: blocks.newPlainTextObject('Link existing Teams chat'),
			}),
			blocks.newButtonElement({
				actionId: `${UIActionId.BridgeLinkChannelClicked}--${room.id}`,
				text: blocks.newPlainTextObject('Link existing Teams channel'),
			}),
		],
	});
	return {
		id: UIElementId.BridgeRoomContextualBarId,
		title: blocks.newPlainTextObject(TITLE),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		blocks: blocks.getBlocks(),
	};
};

// Step 2: pick an existing Teams chat from the ones the bot account belongs to.
export const buildChatPickerViewAsync = async (options: {
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
	modify: IModify;
	app: TeamsBridgeApp;
	room: IRoom;
}): Promise<IUIKitSurfaceViewParam> => {
	const { read, http, persistence, modify, app, room } = options;
	const appUser = await read.getUserReader().getAppUser();
	if (!appUser) {
		return buildBridgeMessageView(modify, ['⚠️ App user not found.']);
	}

	const accessToken = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });
	if (!accessToken) {
		return buildBridgeMessageView(modify, [
			'⚠️ The bot is not logged in to Microsoft Teams.',
			"Use *Teams bot: log in / log out* from the room menu first, then reopen this.",
		]);
	}

	const result = await listMyChatsAsync(http, accessToken);
	const chats = (result?.chats ?? []).filter((c) => c.id);
	if (chats.length === 0) {
		return buildBridgeMessageView(modify, ['ℹ️ No Teams chats were found for the bot account.']);
	}

	const blocks = modify.getCreator().getBlockBuilder();
	blocks.addSectionBlock({ text: blocks.newMarkdownTextObject('*Select the existing Teams chat to link this room to:*') });
	blocks.addActionsBlock({
		blockId: 'TeamsBridge.BridgeChatPicker',
		elements: [
			blocks.newStaticSelectElement({
				actionId: `${UIActionId.BridgeChatSelected}--${room.id}`,
				placeholder: blocks.newPlainTextObject('Select a Teams chat'),
				options: chats.map((c) => ({
					text: blocks.newPlainTextObject(truncate(c.topic || c.id)),
					value: c.id,
				})),
			}),
		],
	});
	return {
		id: UIElementId.BridgeRoomContextualBarId,
		title: blocks.newPlainTextObject(TITLE),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		blocks: blocks.getBlocks(),
	};
};

// Step 2b: pick a channel from the Teams the bot account belongs to.
export const buildChannelPickerViewAsync = async (options: {
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
	modify: IModify;
	app: TeamsBridgeApp;
	room: IRoom;
}): Promise<IUIKitSurfaceViewParam> => {
	const { read, http, persistence, modify, app, room } = options;
	const appUser = await read.getUserReader().getAppUser();
	if (!appUser) {
		return buildBridgeMessageView(modify, ['⚠️ App user not found.']);
	}

	const accessToken = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });
	if (!accessToken) {
		return buildBridgeMessageView(modify, [
			'⚠️ The bot is not logged in to Microsoft Teams.',
			"Use *Teams bot: log in / log out* from the room menu first, then reopen this.",
		]);
	}

	const teams = await listJoinedTeamsAsync(http, accessToken);
	if (teams === null) {
		return buildBridgeMessageView(modify, [
			'⚠️ Could not list the bot account\'s Teams — Microsoft Graph refused the request.',
			'The bot\'s login is probably missing Team/Channel permissions. Log the bot out and back in to grant the new permissions, and ensure the Azure app registration includes: Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMember.Read.All, ChannelMessage.Read.All, ChannelMessage.Send.',
		]);
	}
	if (teams.length === 0) {
		return buildBridgeMessageView(modify, ['ℹ️ The bot account does not belong to any Teams.']);
	}

	const options_: Array<{ label: string; value: string }> = [];
	for (const team of teams) {
		const channels = (await listTeamChannelsAsync(http, team.id, accessToken)) ?? [];
		for (const channel of channels) {
			options_.push({ label: `${team.displayName} › ${channel.displayName}`, value: `${team.id}|${channel.id}` });
		}
	}
	if (options_.length === 0) {
		return buildBridgeMessageView(modify, ['ℹ️ No channels found in the Teams the bot account belongs to.']);
	}

	const blocks = modify.getCreator().getBlockBuilder();
	blocks.addSectionBlock({ text: blocks.newMarkdownTextObject('*Select the Teams channel to link this room to:*') });
	blocks.addSectionBlock({
		text: blocks.newMarkdownTextObject(
			'_Once linked: every channel member must be mapped to a Rocket.Chat user, and members can only be added on the Teams side._',
		),
	});
	blocks.addActionsBlock({
		blockId: 'TeamsBridge.BridgeChannelPicker',
		elements: [
			blocks.newStaticSelectElement({
				actionId: `${UIActionId.BridgeChannelSelected}--${room.id}`,
				placeholder: blocks.newPlainTextObject('Select a Teams channel'),
				options: options_.map((o) => ({ text: blocks.newPlainTextObject(truncate(o.label)), value: o.value })),
			}),
		],
	});
	return {
		id: UIElementId.BridgeRoomContextualBarId,
		title: blocks.newPlainTextObject(TITLE),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		blocks: blocks.getBlocks(),
	};
};

// Step 3: force-map every Teams member of the chosen chat to a Rocket.Chat user.
export const buildMemberMappingViewAsync = async (options: {
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
	modify: IModify;
	app: TeamsBridgeApp;
	room: IRoom;
	threadId: string; // chat thread id, or channel id when teamId is set
	teamId?: string;
}): Promise<IUIKitSurfaceViewParam> => {
	const { read, http, persistence, modify, app, room, threadId, teamId } = options;
	const appUser = await read.getUserReader().getAppUser();
	if (!appUser) {
		return buildBridgeMessageView(modify, ['⚠️ App user not found.']);
	}

	const accessToken = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });
	if (!accessToken) {
		return buildBridgeMessageView(modify, ['⚠️ The bot is not logged in to Microsoft Teams.']);
	}

	const teamsResult = teamId ? await getChannelMembersAsync(http, teamId, threadId, accessToken) : await getTeamsChatMembersAsync(http, threadId, accessToken);
	const teamsMembers = (teamsResult?.members ?? []).filter((m) => m.userId);
	const rcMembers = (await read.getRoomReader().getMembers(room.id)).filter((u) => u.id !== appUser.id && u.type === 'user');

	if (teamsMembers.length === 0) {
		return buildBridgeMessageView(modify, ['⚠️ No members found in that Teams chat.']);
	}
	if (rcMembers.length === 0) {
		return buildBridgeMessageView(modify, [
			'⚠️ This Rocket.Chat room has no non-bot members to map to.',
			'Add the Rocket.Chat users who should represent the Teams members, then try again.',
		]);
	}

	const blocks = modify.getCreator().getBlockBuilder();
	blocks.addSectionBlock({
		text: blocks.newMarkdownTextObject(
			`Map *all ${teamsMembers.length}* Teams member(s) to a Rocket.Chat user. Every field is required — the link is saved only when all are mapped.`,
		),
	});

	const rcOptions = rcMembers.map((u) => ({
		text: blocks.newPlainTextObject(u.name ? `${u.name} (@${u.username})` : `@${u.username}`),
		value: u.id,
	}));

	teamsMembers.forEach((m) => {
		blocks.addInputBlock({
			blockId: `${BridgeMapBlockPrefix}--${m.userId}`,
			optional: false,
			element: blocks.newStaticSelectElement({
				actionId: `${BridgeMapSelectPrefix}--${m.userId}`,
				placeholder: blocks.newPlainTextObject('Select a Rocket.Chat user'),
				options: rcOptions,
			}),
			label: blocks.newPlainTextObject(m.displayName || m.userId),
		});
	});

	return {
		id: UIElementId.BridgeRoomContextualBarId,
		title: blocks.newPlainTextObject(TITLE),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		submit: blocks.newButtonElement({
			// Channel links carry "teamId|channelId" in the middle segment; chat links carry the thread id.
			actionId: `${UIActionId.BridgeLinkSubmit}--${teamId ? `${teamId}|${threadId}` : threadId}--${room.id}`,
			text: blocks.newPlainTextObject('Link & save mappings'),
		}),
		blocks: blocks.getBlocks(),
	};
};

// Entry point: open the bridge chooser bar (or tell the user it is already bridged).
export const openBridgeRoomContextualBarAsync = async (
	triggerId: string,
	room: IRoom,
	operator: IUser,
	read: IRead,
	modify: IModify,
	http: IHttp,
	persistence: IPersistence,
	app: TeamsBridgeApp,
): Promise<void> => {
	const appUser = await read.getUserReader().getAppUser();
	if (!appUser) {
		return;
	}
	// Only block when the room is truly linked to a specific Teams chat. A room can carry a
	// stale isBridged flag with no thread (e.g. after a reinstall) — allow linking in that case.
	const roomRecord = await Room.findByRCRoomId(read, room.id);
	if (roomRecord?.teamsThreadId) {
		// Resolve the friendly chat name when possible, so the notice isn't a raw thread id.
		let chatName = roomRecord.teamsThreadName ?? '';
		if (!chatName) {
			const token = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });
			if (token) {
				const result = await listMyChatsAsync(http, token);
				chatName = result?.chats.find((c) => c.id === roomRecord.teamsThreadId)?.topic ?? '';
			}
		}
		const label = chatName ? `*${chatName}*` : `thread ${roomRecord.teamsThreadId}`;
		await notifyRocketChatUserInRoomAsync(
			`This room is already linked to the Teams chat ${label}. Unlink it first to change the target.`,
			appUser,
			operator,
			room,
			read.getNotifier(),
		);
		return;
	}
	await modify.getUiController().openSurfaceView(buildBridgeChooserView(modify, room), { triggerId }, operator);
};
