import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { DefaultThreadName } from "../Const";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { mapRocketChatMessageToTeamsMessageV2 } from "../MessageHelper";
import {
    createChatThreadAsync,
    createOneOnOneChatThreadAsync,
    sendFileMessageToChatThreadAsync,
    sendTextMessageToChatThreadAsync,
    shareOneDriveFileAsync,
} from "../MicrosoftGraphApi";
import { MessageMapping, OneDriveFile, Room, UserMapping } from "../PersistHelper";
import { PreventRegistry } from "../PreventRegistry";

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
    const appUser = await read.getUserReader().getAppUser(app.getID());
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

    if (!roomRecord.bridgeUserRocketChatUserId) {
        throw new Error("No bridge user assigned to Teams interop room!");
    }

    const bridgeUser = await UserMapping.findByRCUserId(
        read,
        roomRecord.bridgeUserRocketChatUserId
    );
    let userAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: roomRecord.bridgeUserRocketChatUserId,
        app,
        http,
    });
    if (!userAccessToken || !bridgeUser) {
        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
        throw new Error("Invalid bridge user!");
    }

    if (!roomRecord.teamsThreadId) {
        const members = await read.getRoomReader().getMembers(roomId);

        if (
            message.room.type === RoomType.DIRECT_MESSAGE &&
            members.length === 2
        ) {
            const otherMember = members.find((m) => m.id !== bridgeUser.rocketChatUserId);
            if (!otherMember) {
                console.log("Bridge user is sending a message to self, stop processing.");
                return;
            }
            const otherUser = await UserMapping.findByRCUserId(read, otherMember.id);
            if (!otherUser) {
                console.log("Other member has no Teams mapping, stop processing.");
                return;
            }
            const response = await createOneOnOneChatThreadAsync(
                http,
                bridgeUser.teamsUserId,
                otherUser.teamsUserId,
                userAccessToken
            );
            roomRecord.teamsThreadId = response.threadId;
        } else {
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
        }

        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            roomRecord.bridgeUserRocketChatUserId
        );
    }

    let messageText = message.text;
    if (!messageText) {
        messageText = "";
    }

    const isMessageBridged =
        bridgeUser.rocketChatUserId !== message.sender.id;
    let originalSenderName = isMessageBridged
        ? message.sender.name
        : undefined;

    const senderUserAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: message.sender.id,
        app,
        http,
    });
    if (senderUserAccessToken) {
        // If message sender already logged in, make the message sent by themselves instead of via the bridge user
        userAccessToken = senderUserAccessToken;
        originalSenderName = undefined;
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
