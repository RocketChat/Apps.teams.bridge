import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { DefaultThreadName, UnsupportedScenarioHintMessageText } from "../Const";
import { getAppAccessTokenAsync, getUserAccessTokenAsync } from "../AuthHelper";
import { mapRocketChatMessageToTeamsMessageV2 } from "../MessageHelper";
import {
    createChatThreadAsync,
    sendFileMessageToChatThreadAsync,
    sendTextMessageToChatThreadAsync,
    shareOneDriveFileAsync,
} from "../MicrosoftGraphApi";
import { MessageMapping, OneDriveFile, Room, UserMapping } from "../PersistHelper";
import { PreventRegistry } from "../PreventRegistry";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { notifyRocketChatUserInRoomAsync } from "../Notifier";

export const handlePostMessageSentAsync = async (options: {
    message: IMessage;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const { message, read, persistence, app, http } = options;

    if (await PreventRegistry.capture(persistence, `PreventPostMessageHook/${message.id}`)) {
        return;
    }

    // Skip messages relayed from Teams (the app bot is the sender in that case)
    const appUser = await read.getUserReader().getAppUser(app.getID()) as IUser;
    if (appUser && message.sender.id === appUser.id) {
        return;
    }

    const roomId = message.room.id;
    if (!await Room.isBridged(read, roomId)) {
        return;
    }

    const roomRecord = await Room.findByRCRoomId(read, roomId);
    if (!roomRecord) {
        throw new Error("No room record found for Teams interop room!");
    }

    // Determine the access token to use: prefer the sender's own token (logged-in),
    // fall back to the app-level token for non-logged-in users.
    let userAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: message.sender.id,
        app,
        http,
    });
    let originalSenderName: string | undefined;

    if (!userAccessToken) {
        userAccessToken = await getAppAccessTokenAsync({ http, app });
        originalSenderName = message.sender.name;
    }

    if (!userAccessToken) {
        const notifier = read.getNotifier();
        await notifyRocketChatUserInRoomAsync(
            UnsupportedScenarioHintMessageText("No valid access token available"),
            appUser,
            message.sender,
            message.room,
            notifier
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
        const response = await createChatThreadAsync(
            http,
            teamsIds,
            roomName,
            userAccessToken
        );
        roomRecord.teamsThreadId = response.threadId;

        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId
        );
    }

    let messageText = message.text;
    if (!messageText) {
        messageText = "";
    }

    let teamsMessageId = "";
    let rocketChatMessageId = "";
    if (message.file) {
        // If message is a file, use send file operation
        let textMessage = "";
        if (message.attachments && message.attachments[0].description) {
            textMessage = message.attachments[0].description;
        }

        const oneDriveFile = await OneDriveFile.find(
            read,
            message.file.name
        );
        if (!oneDriveFile) {
            return;
        }

        const shareRecord = await shareOneDriveFileAsync(
            http,
            oneDriveFile?.driveItemId,
            userAccessToken
        );

        // Send the message to the chat thread
        const response = await sendFileMessageToChatThreadAsync(
            http,
            textMessage,
            oneDriveFile.fileName,
            shareRecord.shareLink,
            roomRecord.teamsThreadId,
            userAccessToken
        );

        teamsMessageId = response.messageId;
        rocketChatMessageId = message.id as string;
    } else {
        const { text, attachments } = await mapRocketChatMessageToTeamsMessageV2({
            message,
            originalSenderName,
            read,
            http,
            accessToken: userAccessToken,
            messageIdMapping: {
                rocketChatMessageId,
                teamsMessageId,
                teamsThreadId: roomRecord.teamsThreadId,
            }
        });
        messageText = text;

        // Send the message to the chat thread
        const response = await sendTextMessageToChatThreadAsync({
            http,
            textMessage: messageText ?? '',
            threadId: roomRecord.teamsThreadId,
            userAccessToken,
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
    });
};

