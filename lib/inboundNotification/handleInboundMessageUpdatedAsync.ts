import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../../TeamsBridgeApp';
import { mapTeamsMessageToRocketChatMessage } from '../MessageHelper';
import { MessageMapping } from '../PersistHelper';
import { getMessageWithResourceStringAsync } from '../graph';
import type { InBoundNotification } from './handleInboundNotificationAsync';
import { PreventRegistry } from '../PreventRegistry';

export const handleInboundMessageUpdatedAsync = async (
	userAccessToken: string,
	inBoundNotification: InBoundNotification,
	read: IRead,
	modify: IModify,
	http: IHttp,
	persis: IPersistence,
	app: TeamsBridgeApp,
): Promise<void> => {
	const resourceString = inBoundNotification.resourceString;
	const getMessageResponse = await getMessageWithResourceStringAsync(http, resourceString, userAccessToken);

	const messageIdMapping = await MessageMapping.findByTeamsMessageId(read, getMessageResponse.messageId);
	if (!messageIdMapping) {
		// If there's not an existing rocket chat message, stop processing
		return;
	}

	if (await PreventRegistry.capture(persis, `PreventPostMessageUpdateHook/${messageIdMapping.rocketChatMessageId}`)) {
		return;
	}
	const fromUserTeamsId = getMessageResponse.fromTeamsUser.id;
	if (!fromUserTeamsId) {
		// If there's not a sender, stop processing
		return;
	}

	const message = await read.getMessageReader().getById(messageIdMapping.rocketChatMessageId);
	if (!message) {
		// If there's not an existing rocket chat message, stop processing
		return;
	}

	const sender: IUser = message.sender;
	const updatedMessage = await mapTeamsMessageToRocketChatMessage({
		getMessageResponse,
		accessToken: userAccessToken,
		room: message.room,
		sender,
		http,
		modify,
		read,
		uploadFiles: false,
		app,
		persistence: persis,
	});

	const updator = modify.getUpdater();
	let messageBuilder = await updator.message(messageIdMapping.rocketChatMessageId, sender);

	messageBuilder = messageBuilder.setText(updatedMessage.text).setEditor(sender);
	await PreventRegistry.set(persis, `PreventPostMessageUpdateHook/${messageIdMapping.rocketChatMessageId}`);
	await updator.finish(messageBuilder);
};
