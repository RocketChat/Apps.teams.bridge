import type { IHttp, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IMessage } from '@rocket.chat/apps-engine/definition/messages';

import type { TeamsBridgeApp } from '../../TeamsBridgeApp';
import { getUserAccessTokenAsync } from '../AuthHelper';
import { UnsupportedScenarioHintMessageText } from '../Const';
import { mapRocketChatMessageToTeamsMessageV2 } from '../MessageHelper';
import { updateTextMessageInChatThreadAsync } from '../MicrosoftGraphApi';
import { notifyRocketChatUserInRoomAsync } from '../Notifier';
import { MessageMapping } from '../PersistHelper';
import { PreventRegistry } from '../PreventRegistry';

export const handlePostMessageUpdatedAsync = async (options: {
	message: IMessage;
	read: IRead;
	persistence: IPersistence;
	app: TeamsBridgeApp;
	http: IHttp;
}): Promise<void> => {
	const { message, read, persistence, app, http } = options;
	if (!message?.id || !message.text) {
		return;
	}

	if (await PreventRegistry.capture(persistence, `PreventPostMessageUpdateHook/${message.id}`)) {
		return;
	}

	const messageIdMapping = await MessageMapping.findByRCMessageId(read, message.id);
	if (!messageIdMapping) {
		return;
	}

	const senderId = message.sender.id;
	const senderUserAccessToken = await getUserAccessTokenAsync({
		read,
		persistence,
		rocketChatUserId: senderId,
		app,
		http,
	});

	if (senderUserAccessToken) {
		await PreventRegistry.set(persistence, `PreventPostMessageUpdateHook/${message.id}`);
		const { text, attachments } = await mapRocketChatMessageToTeamsMessageV2({
			message,
			read,
			http,
			accessToken: senderUserAccessToken,
			messageIdMapping,
		});
		await updateTextMessageInChatThreadAsync({
			http,
			textMessage: text,
			messageType: 'html',
			messageId: messageIdMapping.teamsMessageId,
			threadId: messageIdMapping.teamsThreadId,
			userAccessToken: senderUserAccessToken,
			attachments,
		});
	} else {
		// Sender is not logged in — use app user token to relay the edit.

		const appUser = await read.getUserReader().getAppUser();
		let appAccessToken: string | null = null;

		if (appUser) {
			appAccessToken = await getUserAccessTokenAsync({
				http,
				app,
				persistence,
				read,
				rocketChatUserId: appUser.id,
			});
		} else {
			return;
		}

		if (!appAccessToken) {
			await notifyRocketChatUserInRoomAsync(
				UnsupportedScenarioHintMessageText('No valid access token available to update message'),
				appUser,
				message.sender,
				message.room,
				read.getNotifier(),
			);
			return;
		}

		await PreventRegistry.set(persistence, `PreventPostMessageUpdateHook/${message.id}`);
		const { text, attachments } = await mapRocketChatMessageToTeamsMessageV2({
			message,
			read,
			originalSenderName: message.sender.name || message.sender.username,
			forceBridgedMessage: true,
			http,
			accessToken: appAccessToken,
			messageIdMapping,
		});
		await updateTextMessageInChatThreadAsync({
			http,
			textMessage: text,
			messageType: 'html',
			messageId: messageIdMapping.teamsMessageId,
			threadId: messageIdMapping.teamsThreadId,
			userAccessToken: appAccessToken,
			attachments,
		});
	}
};
