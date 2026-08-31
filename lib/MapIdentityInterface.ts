import type { IHttp, IModify, IPersistence, IRead, IUIKitSurfaceViewParam } from '@rocket.chat/apps-engine/definition/accessors';
import type { IRoom } from '@rocket.chat/apps-engine/definition/rooms';
import { UIKitSurfaceType } from '@rocket.chat/apps-engine/definition/uikit';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { getUserAccessTokenAsync } from './AuthHelper';
import { UIActionId, UIElementId } from './Const';
import { getChannelMembersAsync, getTeamsChatMembersAsync, getTeamsUserProfileByIdAsync } from './MicrosoftGraphApi';
import { notifyRocketChatUserInRoomAsync } from './Notifier';
import { Room, UserMapping } from './PersistHelper';

export const MapIdentityRcUserBlockId = 'TeamsBridge.MapIdentity.RcUserBlock';
export const MapIdentityTeamsUserBlockId = 'TeamsBridge.MapIdentity.TeamsUserBlock';
export const MapIdentityRcUsernameBlockId = 'TeamsBridge.MapIdentity.RcUsernameBlock';

// Opens the "Map Teams identity" contextual bar: two dropdowns (a room member and a
// member of the linked Teams chat). On submit the two are linked in UserMapping so
// incoming Teams messages from that Teams user post as the chosen Rocket.Chat user.
export const openMapIdentityContextualBarAsync = async (
	triggerId: string,
	currentRoom: IRoom,
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

	const roomRecord = await Room.findByRCRoomId(read, currentRoom.id);
	if (!roomRecord?.teamsThreadId) {
		await notifyRocketChatUserInRoomAsync(
			'This room is not linked to a Teams chat yet. Bridge the room (add the bot) before mapping identities.',
			appUser,
			operator,
			currentRoom,
			read.getNotifier(),
		);
		return;
	}

	const accessToken = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });
	if (!accessToken) {
		await notifyRocketChatUserInRoomAsync(
			'Could not reach Microsoft Teams. Ensure the app bot user is logged in (/teamsbridge-login-app-user).',
			appUser,
			operator,
			currentRoom,
			read.getNotifier(),
		);
		return;
	}

	const teamsResult = roomRecord.teamsTeamId
		? await getChannelMembersAsync(http, roomRecord.teamsTeamId, roomRecord.teamsThreadId, accessToken)
		: await getTeamsChatMembersAsync(http, roomRecord.teamsThreadId, accessToken);
	const teamsMembers = teamsResult?.members ?? [];

	const rcMembers = (await read.getRoomReader().getMembers(currentRoom.id)).filter((u) => u.id !== appUser.id && u.type === 'user');

	if (teamsMembers.length === 0) {
		await notifyRocketChatUserInRoomAsync(
			'Nothing to map: no Teams members found in the linked chat/channel.',
			appUser,
			operator,
			currentRoom,
			read.getNotifier(),
		);
		return;
	}

	// Suggest likely RC matches for unmapped Teams members: try the Teams account's
	// email local-part and name-derived usernames against existing RC usernames.
	const suggestions: Array<{ teamsName: string; rcUsername: string }> = [];
	for (const member of teamsMembers.slice(0, 10)) {
		const existing = await UserMapping.findByTeamsUserId(read, member.userId);
		if (existing) {
			continue;
		}
		const candidates: string[] = [];
		try {
			const profile = await getTeamsUserProfileByIdAsync(http, accessToken, member.userId);
			const local = (profile?.mail ?? '').split('@')[0].toLowerCase();
			if (local) {
				candidates.push(local);
			}
		} catch (error) {
			// no profile — fall back to name-derived candidates
		}
		const name = (member.displayName ?? '').trim().toLowerCase();
		if (name) {
			candidates.push(name.replace(/\s+/g, '.'), name.replace(/\s+/g, ''));
		}
		for (const candidate of [...new Set(candidates)]) {
			if (!candidate) {
				continue;
			}
			try {
				const rcUser = await read.getUserReader().getByUsername(candidate);
				if (rcUser) {
					suggestions.push({ teamsName: member.displayName || member.userId, rcUsername: rcUser.username });
					break;
				}
			} catch (error) {
				// no such user — try next candidate
			}
		}
	}

	const blocks = modify.getCreator().getBlockBuilder();

	if (suggestions.length > 0) {
		blocks.addSectionBlock({
			text: blocks.newMarkdownTextObject(
				`💡 *Suggested matches* — type the username below:\n${suggestions.map((sug) => `• ${sug.teamsName} → \`${sug.rcUsername}\``).join('\n')}`,
			),
		});
	}


	if (rcMembers.length > 0) {
		blocks.addInputBlock({
			blockId: MapIdentityRcUserBlockId,
			optional: true,
			element: blocks.newStaticSelectElement({
				actionId: UIActionId.MapIdentityRcUserSelect,
				placeholder: blocks.newPlainTextObject('Select a Rocket.Chat user'),
				options: rcMembers.map((u) => ({
					text: blocks.newPlainTextObject(u.name ? `${u.name} (@${u.username})` : `@${u.username}`),
					value: u.id,
				})),
			}),
			label: blocks.newPlainTextObject('Rocket.Chat user (room member)'),
		});
	}

	// Any Rocket.Chat user by username — not limited to room members. The mapped user is
	// added to the room automatically on save (mapped users pass the channel join guard).
	blocks.addInputBlock({
		blockId: MapIdentityRcUsernameBlockId,
		optional: true,
		element: blocks.newPlainTextInputElement({
			actionId: UIActionId.MapIdentityRcUsernameInput,
			placeholder: blocks.newPlainTextObject('e.g. bridgetest1'),
		}),
		label: blocks.newPlainTextObject('…or type any Rocket.Chat username'),
	});

	blocks.addInputBlock({
		blockId: MapIdentityTeamsUserBlockId,
		element: blocks.newStaticSelectElement({
			actionId: UIActionId.MapIdentityTeamsUserSelect,
			placeholder: blocks.newPlainTextObject('Select a Teams user'),
			options: teamsMembers.map((m) => ({
				text: blocks.newPlainTextObject(m.displayName || m.userId),
				value: m.userId,
			})),
		}),
		label: blocks.newPlainTextObject('Teams user (from the linked chat)'),
	});

	const view: IUIKitSurfaceViewParam = {
		id: UIElementId.MapIdentityContextualBarId,
		title: blocks.newPlainTextObject('Map Teams identity'),
		type: UIKitSurfaceType.CONTEXTUAL_BAR,
		submit: blocks.newButtonElement({
			actionId: `${UIActionId.MapIdentitySubmit}--${currentRoom.id}`,
			text: blocks.newPlainTextObject('Save mapping'),
		}),
		blocks: blocks.getBlocks(),
	};

	await modify.getUiController().openSurfaceView(view, { triggerId }, operator);
};
