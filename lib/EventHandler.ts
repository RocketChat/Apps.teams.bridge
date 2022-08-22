import {
    IHttp,
    IModify,
    IPersistence,
    IRead
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom, RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import {
    AuthenticationEndpointPath,
    LoginRequiredHintMessageText,
    UnsupportedMessageTypeHintMessageText
} from "./Const";
import {
    generateHintMessageWithTeamsLoginButton,
    notifyRocketChatUserAsync,
    notifyRocketChatUserInRoomAsync
} from "./Messages";
import {
    createOneOnOneChatThreadAsync,
    deleteTextMessageInChatThreadAsync,
    sendTextMessageToChatThreadAsync,
    updateTextMessageInChatThreadAsync
} from "./MicrosoftGraphApi";
import {
    checkDummyUserByRocketChatUserIdAsync,
    persistMessageIdMappingAsync,
    retrieveDummyUserByRocketChatUserIdAsync,
    retrieveMessageIdMappingByRocketChatMessageIdAsync,
    retrieveUserAccessTokenAsync,
    retrieveUserByRocketChatUserIdAsync
} from "./PersistHelper";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "./UrlHelper";
import { shortnameToUnicode } from "emojione";

export const handlePreMessageSentPreventAsync = async (message: IMessage, read: IRead) : Promise<boolean> => {
    if (message.threadId) {
        const isTeamsMessageThread = await isTeamsMessageAsync(message.threadId, read);
        if (isTeamsMessageThread) {
            // There's no thread message concept in Teams

            const appUser = (await read.getUserReader().getAppUser()) as IUser;
            const notifier = read.getNotifier();

            await notifyRocketChatUserInRoomAsync(
                UnsupportedMessageTypeHintMessageText("Thread Message"),
                appUser,
                message.sender,
                message.room,
                notifier);

            return true;
        }
    }

    return false;
};

export const handlePostMessageSentAsync = async (
    message: IMessage,
    read: IRead,
    http: IHttp,
    persistence: IPersistence,
    modify: IModify,
    app: TeamsBridgeApp) : Promise<void> => {
    // In first version, we'll only support sending text message in 1:1 chat with Teams user

    // If the message is not a text message, stop processing
    if (!message || !message.text || !isTextMessage(message)) {
        return;
    }

    // If the message is sent in a thread, stop processing
    if (message.threadId) {
        return;
    }

    // If the message is not sent in a 1:1 chat, stop processing
    const room = message.room;
    if (!isOneOnOneChat(room)) {
        return;
    }

    // If the message receiver is not a Teams Dummy user, stop processing
    const receiverId = getDirectMessageReceiverId(message, room);
    if (receiverId === undefined) {
        return;
    }
    const isReceiverDummyUser = await checkDummyUserByRocketChatUserIdAsync(read, receiverId);
    if (!isReceiverDummyUser) {
        return;
    }

    // If the message sender has not logged in to Teams
    // Send a notification to let the sender know he need to logged in to Teams to start cross platform collaboration
    const senderId = message.sender.id;
    const senderUserAccessToken = await retrieveUserAccessTokenAsync(read, senderId);
    if (!senderUserAccessToken) {
        await notifyNotLoggedInUserAsync(read, message.sender, room, modify, app, LoginRequiredHintMessageText);
        return;
    }

    // All checks passed, find the dummy user record
    const dummyUser = await retrieveDummyUserByRocketChatUserIdAsync(read, receiverId);
    const senderUser = await retrieveUserByRocketChatUserIdAsync(read, senderId);

    if (!dummyUser || !senderUser) {
        // If any of dummy user or sender user information is missing, stop processing
        return;
    }

    // Create a 1 on 1 chat thread in Teams
    // TODO: find existing Teams thread instead of create a new one
    const teamsThread = await createOneOnOneChatThreadAsync(
        http,
        senderUser.teamsUserId,
        dummyUser.teamsUserId,
        senderUserAccessToken);

    // Handle emoji in text
    const messageText = shortnameToUnicode(message.text);

    // Send the message to the chat thread
    const response = await sendTextMessageToChatThreadAsync(
        http,
        messageText,
        teamsThread.threadId,
        senderUserAccessToken);

    const teamsMessageId = response.messageId;
    const rocketChatMessageId = message.id as string;

    await persistMessageIdMappingAsync(persistence, rocketChatMessageId, teamsMessageId, teamsThread.threadId);
};

export const handlePreMessageOperationPreventAsync = async (message: IMessage, read: IRead): Promise<boolean> => {
    const isTeamsMessage = await isTeamsMessageAsync(message.id, read);
    if (!isTeamsMessage) {
        return false;
    }

    // If the user that operate the Teams message has not logged in to Teams
    // Send a notification to let the sender know he need to logged in to Teams to apply the operation
    const senderId = message.sender.id;
    const senderUserAccessToken = await retrieveUserAccessTokenAsync(read, senderId);

    if (!senderUserAccessToken) {
        return true;
    }

    return false;
}

export const handlePostMessageUpdatedAsync = async (message: IMessage, read: IRead, http: IHttp): Promise<void> => {
    if (!message || !message.id || !message.text) {
        return;
    }

    const messageIdMapping = await retrieveMessageIdMappingByRocketChatMessageIdAsync(read, message.id);
    if (!messageIdMapping) {
        return;
    }

    const senderId = message.sender.id;
    const senderUserAccessToken = await retrieveUserAccessTokenAsync(read, senderId);
    if (!senderUserAccessToken) {
        return;
    }

    await updateTextMessageInChatThreadAsync(
        http,
        message.text,
        messageIdMapping.teamsMessageId,
        messageIdMapping.teamsThreadId,
        senderUserAccessToken);
};

export const handlePostMessageDeletedAsync = async (message: IMessage, read: IRead, http: IHttp): Promise<void> => {
    if (!message || !message.id || !message.text) {
        return;
    }

    const messageIdMapping = await retrieveMessageIdMappingByRocketChatMessageIdAsync(read, message.id);
    if (!messageIdMapping) {
        return;
    }

    const senderId = message.sender.id;
    const senderUserAccessToken = await retrieveUserAccessTokenAsync(read, senderId);
    if (!senderUserAccessToken) {
        return;
    }
    
    const senderUser = await retrieveUserByRocketChatUserIdAsync(read, senderId);
    if (!senderUser) {
        return;
    }

    await deleteTextMessageInChatThreadAsync(
        http,
        senderUser.teamsUserId,
        messageIdMapping.teamsMessageId,
        messageIdMapping.teamsThreadId,
        senderUserAccessToken);
};

const isTeamsMessageAsync = async (messageId: string | undefined, read: IRead) : Promise<boolean> => {
    if (!messageId) {
        return false;
    }

    const messageIdMapping = await retrieveMessageIdMappingByRocketChatMessageIdAsync(read, messageId);
    if (messageIdMapping) {
        return true;
    }

    return false;
};

const notifyNotLoggedInUserAsync = async (
    read: IRead,
    user: IUser,
    room: IRoom,
    modify: IModify,
    app: TeamsBridgeApp,
    hintMessageText: string) : Promise<void> => {
    const appUser = (await read.getUserReader().getAppUser()) as IUser;

    const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
    const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
    const accessors = app.getAccessors();
    const authEndpointUrl = await getRocketChatAppEndpointUrl(accessors, AuthenticationEndpointPath);
    const loginUrl = getLoginUrl(aadTenantId, aadClientId, authEndpointUrl, user.id);
    const message = generateHintMessageWithTeamsLoginButton(loginUrl, appUser, room, hintMessageText)
    
    await notifyRocketChatUserAsync(message, user, modify.getNotifier());
};

const getDirectMessageReceiverId = (message: IMessage, room: IRoom) : string | undefined => {
    return room.userIds?.find(id => id !== message.sender.id);
};

const isOneOnOneChat = (room: IRoom) : boolean => {
    if (room.type === RoomType.DIRECT_MESSAGE && room.userIds && room.userIds.length === 2) {
        return true;
    }

    return false;
};

const isTextMessage = (message: IMessage) : boolean => {
    if (message.text) {
        return true;
    }

    return false;
};
