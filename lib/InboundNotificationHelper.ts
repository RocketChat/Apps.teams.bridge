import {
    IHttp,
    IModify,
    IPersistence,
    IRead
} from "@rocket.chat/apps-engine/definition/accessors";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { sendRocketChatOneOnOneMessageAsync } from "./Messages";
import { getMessageWithResourceStringAsync, MessageContentType } from "./MicrosoftGraphApi";
import {
    persistMessageIdMappingAsync,
    retrieveDummyUserByTeamsUserIdAsync,
    retrieveUserAccessTokenAsync,
    retrieveUserByTeamsUserIdAsync
} from "./PersistHelper";

export enum NotificationChangeType {
    Created = 'created',
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

    const resourceString = inBoundNotification.resourceString;
    const getMessageResponse = await getMessageWithResourceStringAsync(http, resourceString, userAccessToken);

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

    const fromDummyUser = await retrieveDummyUserByTeamsUserIdAsync(read, fromUserTeamsId);
    if (!fromDummyUser) {
        // If there's not a dummy user created for message sender, stop processing
        // TODO: create dummy user on demand after find out how to create user with Rocket.Chat Apps Engine.
        return;
    }

    const receiver : IUser = await read.getUserReader().getById(receiverRocketChatUserId);
    const sender: IUser = await read.getUserReader().getById(fromDummyUser.rocketChatUserId);
    if (!receiver || !sender) {
        // If receiver or sender Rocket.Chat user does not exist, stop processing
        return;
    }

    let messageContent = getMessageResponse.messageContent;
    if (getMessageResponse.messageContentType && getMessageResponse.messageContentType === MessageContentType.Html) {
        // TODO: find a better way to trim html tag from html messages
        messageContent = getMessageResponse.messageContent.replace(/<\/?[^>]+(>|$)/g, "");
    }

    const rocketChatMessageId = await sendRocketChatOneOnOneMessageAsync(messageContent, sender, receiver, read, modify);
    await persistMessageIdMappingAsync(persis, rocketChatMessageId, getMessageResponse.messageId, getMessageResponse.threadId);
};
