import {
    IHttp,
    IModify,
    IPersistence,
    IRead
} from "@rocket.chat/apps-engine/definition/accessors";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { mapTeamsMessageToRocketChatMessage, sendRocketChatMessageInRoomAsync, sendRocketChatOneOnOneMessageAsync } from "./MessageHelper";
import { getMessageWithResourceStringAsync, MessageContentType } from "./MicrosoftGraphApi";
import {
    persistMessageIdMappingAsync,
    retrieveDummyUserByTeamsUserIdAsync,
    retrieveMessageIdMappingByTeamsMessageIdAsync,
    retrieveRoomByTeamsThreadIdAsync,
    retrieveUserAccessTokenAsync,
    retrieveUserByTeamsUserIdAsync
} from "./PersistHelper";

export enum NotificationChangeType {
    Created = 'created',
    Updated = 'updated',
    // Deleted = 'deleted',
};

export enum NotificationResourceType {
    ChatMessage = 'chatMessage',
};

export interface InBoundNotification {
    receiverRocketChatUserId: string,
    subscriptionId: string,
    changeType: NotificationChangeType,
    resourceId: string,
    resourceString: string,
    resourceType: NotificationResourceType,
};

export const handleInboundNotificationAsync = async (
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence) : Promise<void> => {
    console.log('Processing inbound notification!');
    
    const receiverRocketChatUserId = inBoundNotification.receiverRocketChatUserId;
    if (!receiverRocketChatUserId) {
        // If there's not a receiver, stop processing
        return;
    }

    const userAccessToken = await retrieveUserAccessTokenAsync(read, receiverRocketChatUserId);
    if (!userAccessToken) {
        // If receiver's access token does not exist in persist or expired, stop processing
        // TODO: handle this issue when token auto refresh is enabled.
        return;
    }

    switch (inBoundNotification.changeType) {
        case NotificationChangeType.Created:
            await handleInboundMessageCreatedAsync(userAccessToken, inBoundNotification, read, modify, http, persis);
            break;
        // Update / delete message is not supported for Rocket.Chat App
        
        case NotificationChangeType.Updated:
            await handleInboundMessageUpdatedAsync(userAccessToken, inBoundNotification, read, modify, http, persis);
            break;
            /*
        case NotificationChangeType.Deleted:
            await handleInboundMessageDeletedAsync(inBoundNotification, read, modify, http, persis);
            break;
        */
        default:
            console.log(`Unsupported notification change type`);
            return;
    }
};

const handleInboundMessageCreatedAsync = async (
    userAccessToken: string,
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence) : Promise<void> => {
    console.log('Processing message created!');

    const receiverRocketChatUserId = inBoundNotification.receiverRocketChatUserId;
    const resourceString = inBoundNotification.resourceString;
    const getMessageResponse = await getMessageWithResourceStringAsync(http, resourceString, userAccessToken);

    const fromUserTeamsId = getMessageResponse.fromUserTeamsId;
    if (!fromUserTeamsId) {
        // If there's not a sender, stop processing
        return;
    }

    const roomRecord = await retrieveRoomByTeamsThreadIdAsync(read, getMessageResponse.threadId);
    if (!roomRecord) {
        // TODO: handle thread created in Teams scenario
        return;
    }

    if (roomRecord.bridgeUserRocketChatUserId && roomRecord.bridgeUserRocketChatUserId === receiverRocketChatUserId) {
        // Only handle notification received by the bridge user to avoid duplication

        const fromDummyUser = await retrieveDummyUserByTeamsUserIdAsync(read, fromUserTeamsId);
        if (!fromDummyUser) {
            // If the message if not from a dummy user, stop processing
            console.log("Message not from dummy user!");
            // TODO: create dummy user on demand.
            // There could be dummy user out of sync issue.
            // If the dummy user has not been created for a recently added Teams user, we need to detect this and create dummy user.
            return;
        }

        const sender: IUser = await read.getUserReader().getById(fromDummyUser.rocketChatUserId);

        const messageContent = mapTeamsMessageToRocketChatMessage(
            getMessageResponse.messageContent,
            getMessageResponse.messageContentType);

        const room = await read.getRoomReader().getById(roomRecord.rocketChatRoomId);
        if (!room) {
            return;
        }
        
        const rocketChatMessageId = await sendRocketChatMessageInRoomAsync(messageContent, sender, room, modify);
        await persistMessageIdMappingAsync(persis, rocketChatMessageId, getMessageResponse.messageId, getMessageResponse.threadId);
    }
};

const handleInboundMessageUpdatedAsync = async (
    userAccessToken: string,
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence) : Promise<void> => {
    console.log('Processing message updated!');

    const receiverRocketChatUserId = inBoundNotification.receiverRocketChatUserId;

    const resourceString = inBoundNotification.resourceString;
    const getMessageResponse = await getMessageWithResourceStringAsync(http, resourceString, userAccessToken);

    const messageIdMapping = await retrieveMessageIdMappingByTeamsMessageIdAsync(read, getMessageResponse.messageId);
    if (!messageIdMapping) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }

    const fromUserTeamsId = getMessageResponse.fromUserTeamsId;
    if (!fromUserTeamsId) {
        // If there's not a sender, stop processing
        return;
    }

    const fromUser = await retrieveUserByTeamsUserIdAsync(read, fromUserTeamsId);
    if (fromUser && fromUser.rocketChatUserId === receiverRocketChatUserId) {
        console.log("This is a notification for message sent by sender himself, skip!");
        return;
    }

    const message = await read.getMessageReader().getById(messageIdMapping.rocketChatMessageId);
    console.log("message to update:");
    console.log(message);
    if (!message) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }

    const sender: IUser = message.sender;
    const messageContent = mapTeamsMessageToRocketChatMessage(
        getMessageResponse.messageContent,
        getMessageResponse.messageContentType);

    const updator = modify.getUpdater();
    let messageBuilder = await updator.message(messageIdMapping.rocketChatMessageId, sender);
    messageBuilder = messageBuilder
        .setText(messageContent)
        .setEditor(sender);
    await updator.finish(messageBuilder);
};

const handleInboundMessageDeletedAsync = async (
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence) : Promise<void> => {
    console.log('Processing message deleted!');
};
