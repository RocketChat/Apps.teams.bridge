import type { IHttp, IPersistence, IRead, IModify } from '@rocket.chat/apps-engine/definition/accessors';
import type { IMessage } from '@rocket.chat/apps-engine/definition/messages';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../../TeamsBridgeApp';
import { getUserAccessTokenAsync } from '../AuthHelper';
import { DefaultThreadName, UnsupportedScenarioHintMessageText } from '../Const';
import { mapRocketChatMessageToTeamsMessageV2 } from '../MessageHelper';
import { createChatThreadAsync, sendFileMessageToChatThreadAsync, sendTextMessageToChatThreadAsync, shareOneDriveFileAsync } from '../MicrosoftGraphApi';
import { notifyRocketChatUserInRoomAsync, notifyRoomMembersAppUserNotLoggedInAsync } from '../Notifier';
import { MessageMapping, OneDriveFile, Room, UserMapping, AppUserLoginNotified, RecentActivity } from '../PersistHelper';
import { PreventRegistry } from '../PreventRegistry';

export const handlePostMessageSentAsync = async (options: {
	message: IMessage;
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
	modify: IModify;
	app: TeamsBridgeApp;
}): Promise<void> => {
	const { message, read, persistence, app, http, modify } = options;

	if (await PreventRegistry.capture(persistence, `PreventPostMessageHook/${message.id}`)) {
		return;
	}

	// Skip messages relayed from Teams (the app bot is the sender in that case)
	const appUser = (await read.getUserReader().getAppUser(app.getID())) as IUser;
	if (message.sender.id === appUser?.id) {
		return;
	}

	const roomId = message.room.id;
	const roomRecord = await Room.findByRCRoomId(read, roomId);
	if (!roomRecord?.isBridged) {
		app.getLogger().debug(`Room ${roomId} is not bridged, skipping message processing.`);
		return;
	}

	let accessToken = await getUserAccessTokenAsync({
		read,
		persistence,
		rocketChatUserId: message.sender.id,
		app,
		http,
	});

	const userHasAccessToken = typeof accessToken === 'string' && accessToken.length > 0;

	if (!userHasAccessToken) {
		const appUserToken = await getUserAccessTokenAsync({
			read,
			persistence,
			rocketChatUserId: appUser.id,
			app,
			http,
		});

		if (typeof appUserToken === 'string' && appUserToken.length > 0) {
			accessToken = appUserToken;
		} else {
			const alreadyNotified = await AppUserLoginNotified.isSetToday(read.getPersistenceReader(), roomId);

			if (alreadyNotified) {
				return;
			}

			await notifyRoomMembersAppUserNotLoggedInAsync({
				read,
				modify,
				http,
				persistence,
				app,
				roomId,
			});
			return;
		}
	}

	if (!accessToken) {
		const notifier = read.getNotifier();
		await notifyRocketChatUserInRoomAsync(
			UnsupportedScenarioHintMessageText('No valid access token available'),
			appUser,
			message.sender,
			message.room,
			notifier,
		);
		return;
	}

	if (!roomRecord.teamsThreadId) {
		const members = await read.getRoomReader().getMembers(roomId);

		const teamsIds: string[] = [];
		for (const member of members) {
			const user = await UserMapping.findByRCUserId(read, member.id);
			if (user) {
				teamsIds.push(user.teamsUserId);
			}
		}

		const roomName = message.room.displayName ?? DefaultThreadName;
		const response = await createChatThreadAsync(http, teamsIds, roomName, accessToken);
		roomRecord.teamsThreadId = response.threadId;

		await Room.persist(read, persistence, roomRecord.rocketChatRoomId, roomRecord.teamsThreadId);
	}

	let messageText = message.text;
	if (!messageText) {
		messageText = '';
	}

	let teamsMessageId = '';
	let rocketChatMessageId = '';

	// Stamp recent activity before sending to Teams
	await RecentActivity.set({
		read,
		persistence,
		rcUserId: userHasAccessToken ? message.sender.id : appUser.id,
		teamsThreadId: roomRecord.teamsThreadId,
		kind: 'create',
	});

	if (message.file) {
		// If message is a file, use send file operation
		let textMessage = '';
		if (message.attachments?.[0].description) {
			textMessage = message.attachments[0].description;
		}

		const oneDriveFile = await OneDriveFile.find(read, message.file.name);
		if (!oneDriveFile) {
			return;
		}

		const shareRecord = await shareOneDriveFileAsync(http, oneDriveFile?.driveItemId, accessToken);

		await PreventRegistry.set(persistence, `PreventPostMessageUpdateHook/${message.id}`);
		const response = await sendFileMessageToChatThreadAsync(
			http,
			textMessage,
			oneDriveFile.fileName,
			shareRecord.shareLink,
			roomRecord.teamsThreadId,
			accessToken,
		);

		teamsMessageId = response.messageId;
		rocketChatMessageId = message.id as string;
	} else {
		const { text, attachments } = await mapRocketChatMessageToTeamsMessageV2({
			message,
			originalSenderName: message.sender.name ? `${message.sender.name} (@${message.sender.username})` : message.sender.username,
			read,
			http,
			accessToken,
			messageIdMapping: {
				rocketChatMessageId,
				teamsMessageId,
				teamsThreadId: roomRecord.teamsThreadId,
			},
			forceBridgedMessage: !userHasAccessToken,
		});
		messageText = text;

		// Send the message to the chat thread
		const response = await sendTextMessageToChatThreadAsync({
			http,
			textMessage: messageText ?? '',
			threadId: roomRecord.teamsThreadId,
			accessToken,
			attachments,
		});

		teamsMessageId = response.messageId;
		rocketChatMessageId = message.id as string;
	}

	await MessageMapping.persist({
		persistence,
		rocketChatMessageId,
		teamsMessageId,
		teamsThreadId: roomRecord.teamsThreadId,
		relayedByAppUser: !userHasAccessToken,
	});
};
