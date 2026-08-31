import type { IHttp, IModify, IPersistence, IRead, IUIKitSurfaceViewParam } from '@rocket.chat/apps-engine/definition/accessors';
import type { IRoom } from '@rocket.chat/apps-engine/definition/rooms';
import { ButtonStyle, UIKitSurfaceType } from '@rocket.chat/apps-engine/definition/uikit';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { isAppUserLoggedInAsync } from './AppUserSession';
import { getUserAccessTokenAsync } from './AuthHelper';
import { UIActionId, UIElementId } from './Const';
import { listMyChatsAsync } from './MicrosoftGraphApi';
import { Room } from './PersistHelper';

// Single entry point for the app: one "Teams Bridge" room-action button opens this hub,
// and every capability is a button inside it (keeps the room kebab menu uncluttered).
export const buildHubViewAsync = async (options: {
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
	modify: IModify;
	app: TeamsBridgeApp;
	room: IRoom;
}): Promise<IUIKitSurfaceViewParam> => {
	const { read, http, persistence, modify, app, room } = options;

	const loggedIn = await isAppUserLoggedInAsync({ read, http, persistence, app });

	// Resolve what the room is mapped to (the linked Teams chat/channel), if anything.
	const roomRecord = await Room.findByRCRoomId(read, room.id);
	const threadId = roomRecord?.teamsThreadId;
	const isMapped = Boolean(threadId);
	let bridgedTo = '';
	if (isMapped) {
		if (roomRecord?.teamsThreadName) {
			bridgedTo = roomRecord.teamsThreadName;
		} else if (loggedIn) {
			const appUser = await read.getUserReader().getAppUser();
			const token = appUser ? await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id }) : undefined;
			if (token) {
				const result = await listMyChatsAsync(http, token);
				const match = result?.chats.find((c) => c.id === threadId);
				if (match?.topic) {
					bridgedTo = match.topic;
				}
			}
		}
		if (!bridgedTo) {
			bridgedTo = `chat ${(threadId as string).slice(0, 14)}…`;
		}
	}

	const blocks = modify.getCreator().getBlockBuilder();

	blocks.addSectionBlock({ text: blocks.newMarkdownTextObject('*Microsoft Teams Bridge*') });
	blocks.addContextBlock({
		elements: [
			blocks.newMarkdownTextObject(
				`Room: ${isMapped ? `🔗 Mapped → *${bridgedTo}*` : '❌ *Not mapped to Teams*'}  •  Bot: ${loggedIn ? '✅ logged in' : '⛔ logged out'}`,
			),
		],
	});

	// Connection controls — always available. Login/logout label reflects live status.
	blocks.addSectionBlock({ text: blocks.newMarkdownTextObject('*Connection*') });
	blocks.addActionsBlock({
		blockId: 'TeamsBridge.Hub.Connection',
		elements: [
			blocks.newButtonElement({
				actionId: `${UIActionId.AppUserSessionButtonClicked}--${room.id}`,
				text: blocks.newPlainTextObject(loggedIn ? 'Log out Teams bot' : 'Log in Teams bot'),
			}),
			blocks.newButtonElement({
				actionId: `${UIActionId.TeamsBridgeHubRefresh}--${room.id}`,
				text: blocks.newPlainTextObject('↻ Refresh'),
			}),
			blocks.newButtonElement({
				actionId: `${UIActionId.TeamsBridgeRestoreMappings}--${room.id}`,
				text: blocks.newPlainTextObject('Restore mappings'),
			}),
		],
	});

	// Member tools — only meaningful once the room is mapped.
	blocks.addSectionBlock({ text: blocks.newMarkdownTextObject('*Members*') });
	blocks.addActionsBlock({
		blockId: 'TeamsBridge.Hub.Members',
		elements: [
			blocks.newButtonElement({
				actionId: `${UIActionId.AddTeamsUserButtonClicked}--${room.id}`,
				text: blocks.newPlainTextObject('Add Teams user'),
			}),
			blocks.newButtonElement({
				actionId: `${UIActionId.ViewTeamsMembersButtonClicked}--${room.id}`,
				text: blocks.newPlainTextObject('View Teams members'),
			}),
			blocks.newButtonElement({
				actionId: `${UIActionId.MapIdentityButtonClicked}--${room.id}`,
				text: blocks.newPlainTextObject('Map Teams identity'),
			}),
		],
	});

	// Primary map / unmap action, always at the bottom and unmissable.
	blocks.addDividerBlock();
	if (isMapped) {
		blocks.addSectionBlock({ text: blocks.newMarkdownTextObject(`This room relays with *${bridgedTo}*.`) });
		blocks.addActionsBlock({
			blockId: 'TeamsBridge.Hub.MapAction',
			elements: [
				blocks.newButtonElement({
					actionId: `${UIActionId.BridgeUnlinkClicked}--${room.id}`,
					text: blocks.newPlainTextObject('Break Teams map'),
					style: ButtonStyle.DANGER,
				}),
			],
		});
	} else {
		blocks.addSectionBlock({
			text: blocks.newMarkdownTextObject('❌ *This room is not mapped to Teams.* Messages are not relayed anywhere.'),
		});
		blocks.addActionsBlock({
			blockId: 'TeamsBridge.Hub.MapAction',
			elements: [
				blocks.newButtonElement({
					actionId: `${UIActionId.BridgeRoomButtonClicked}--${room.id}`,
					text: blocks.newPlainTextObject('Map to Teams'),
					style: ButtonStyle.PRIMARY,
				}),
			],
		});
	}

	return {
		id: UIElementId.TeamsBridgeHubContextualBarId,
		title: blocks.newPlainTextObject('Teams Bridge'),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		blocks: blocks.getBlocks(),
	};
};

// Confirmation step before breaking a Teams map (destructive: relay stops).
export const buildUnlinkConfirmView = (modify: IModify, room: IRoom, mappedName: string): IUIKitSurfaceViewParam => {
	const blocks = modify.getCreator().getBlockBuilder();
	blocks.addSectionBlock({
		text: blocks.newMarkdownTextObject(
			`⚠️ *Break the map between this room and ${mappedName ? `*${mappedName}*` : 'its Teams chat'}?*\nMessages will stop relaying in both directions. User identity mappings are kept.`,
		),
	});
	blocks.addActionsBlock({
		blockId: 'TeamsBridge.Hub.UnlinkConfirm',
		elements: [
			blocks.newButtonElement({
				actionId: `${UIActionId.BridgeUnlinkConfirm}--${room.id}`,
				text: blocks.newPlainTextObject('Yes, break the map'),
				style: ButtonStyle.DANGER,
			}),
			blocks.newButtonElement({
				actionId: `${UIActionId.BridgeUnlinkCancel}--${room.id}`,
				text: blocks.newPlainTextObject('Cancel'),
			}),
		],
	});
	return {
		id: UIElementId.TeamsBridgeHubContextualBarId,
		title: blocks.newPlainTextObject('Teams Bridge'),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		blocks: blocks.getBlocks(),
	};
};

export const openHubViewAsync = async (
	triggerId: string,
	room: IRoom,
	operator: IUser,
	read: IRead,
	http: IHttp,
	persistence: IPersistence,
	modify: IModify,
	app: TeamsBridgeApp,
): Promise<void> => {
	const view = await buildHubViewAsync({ read, http, persistence, modify, app, room });
	await modify.getUiController().openSurfaceView(view, { triggerId }, operator);
};
