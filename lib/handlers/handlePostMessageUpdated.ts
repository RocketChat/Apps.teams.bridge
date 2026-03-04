import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import {
    LoggedInBridgeUserRequiredHintMessageText,
    UnsupportedScenarioHintMessageText,
} from "../Const";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { mapRocketChatMessageToTeamsMessageV2 } from "../MessageHelper";
import { updateTextMessageInChatThreadAsync } from "../MicrosoftGraphApi";
import { notifyNotLoggedInUserAsync, notifyRocketChatUserInRoomAsync } from "../Notifier";
import { MessageMapping, Room } from "../PersistHelper";
import { PreventRegistry } from "../PreventRegistry";

export const handlePostMessageUpdatedAsync = async (options: {
    message: IMessage;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp;
}): Promise<void> => {
    const { message, read, persistence, app, http } = options;
    if (!message || !message.id || !message.text) {
        return;
    }

    if (
        await PreventRegistry.capture(
            persistence,
            `PreventPostMessageUpdateHook/${message.id}`
        )
    ) {
        return;
    }

    const messageIdMapping =
        await MessageMapping.findByRCMessageId(
            read,
            message.id
        );
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
        await PreventRegistry.set(
            persistence,
            `PreventPostMessageUpdateHook/${message.id}`
        );
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
        const bridgeRoom = await Room.findByTeamsThreadId(
            read,
            messageIdMapping.teamsThreadId
        );

        if (bridgeRoom?.bridgeUserRocketChatUserId) {
            const bridgeUserAccessToken = await getUserAccessTokenAsync({
                app,
                http,
                persistence,
                read,
                rocketChatUserId: bridgeRoom.bridgeUserRocketChatUserId,
            });

            if (!bridgeUserAccessToken) {
                const appUser = await read.getUserReader().getAppUser();
                if (appUser) {
                    await notifyRocketChatUserInRoomAsync(
                        UnsupportedScenarioHintMessageText('The session of the bridge user is not valid. Please ask the user to log in again. Messaging without a valid bridge user session'),
                        appUser,
                        message.sender,
                        message.room,
                        read.getNotifier()
                    );
                }
                return;
            }

            await PreventRegistry.set(
                persistence,
                `PreventPostMessageUpdateHook/${message.id}`
            );
            const { text, attachments } = await mapRocketChatMessageToTeamsMessageV2({
                message,
                read,
                originalSenderName: message.sender.name || message.sender.username,
                forceBridgedMessage: true,
                http,
                accessToken: bridgeUserAccessToken,
                messageIdMapping,
            });
            await updateTextMessageInChatThreadAsync({
                http,
                textMessage: text,
                messageType: 'html',
                messageId: messageIdMapping.teamsMessageId,
                threadId: messageIdMapping.teamsThreadId,
                userAccessToken: bridgeUserAccessToken,
                attachments,
            });

        } else {
            notifyNotLoggedInUserAsync(
                read,
                message.sender,
                message.room,
                app,
                LoggedInBridgeUserRequiredHintMessageText
            );
        }
    }
};
