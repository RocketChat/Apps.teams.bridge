import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import { MessageMapping } from '../PersistHelper';
import type { InBoundNotification } from './handleInboundNotificationAsync';
import { PreventRegistry } from '../PreventRegistry';

export const handleInboundMessageDeletedAsync = async (
	inBoundNotification: InBoundNotification,
	read: IRead,
	modify: IModify,
	http: IHttp,
	persis: IPersistence,
): Promise<void> => {
	const resourceString = inBoundNotification.resourceId;

	const messageIdMapping = await MessageMapping.findByTeamsMessageId(read, resourceString);

	if (!messageIdMapping) {
		// If there's not an existing rocket chat message, stop processing
		return;
	}

	if (await PreventRegistry.capture(persis, `PreventPostMessageDeleteHook/${messageIdMapping.rocketChatMessageId}`)) {
		// Prevent duplicate processing
		return;
	}

	const message = await read.getMessageReader().getById(messageIdMapping.rocketChatMessageId);
	if (!message) {
		// If there's not an existing rocket chat message, stop processing
		return;
	}

	const sender: IUser = message.sender;

	const updator = modify.getUpdater();
	let messageBuilder = await updator.message(messageIdMapping.rocketChatMessageId, sender);

	await PreventRegistry.set(persis, `PreventPostMessageUpdateHook/${messageIdMapping.rocketChatMessageId}`);
	messageBuilder = messageBuilder.setText('~This message has been deleted.~').setEditor(sender);
	await updator.finish(messageBuilder);
};
