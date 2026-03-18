import { IUser } from "@rocket.chat/apps-engine/definition/users";
import {
    getMessageWithResourceStringAsync,
    getTeamsUserProfileByIdAsync,
} from "../graph";
import {
    formatTeamsSenderInfo,
    mapTeamsMessageToRocketChatMessage,
} from "../MessageHelper";
import {
    MessageMapping,
    UserMapping,
} from "../PersistHelper";
import { InBoundNotification } from "./handleInboundNotificationAsync";
import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { PreventRegistry } from "../PreventRegistry";
import { getUserAccessTokenAsync } from "../AuthHelper";

export const handleInboundMessageUpdatedAsync = async (
    userAccessToken: string,
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence,
    app: TeamsBridgeApp,
): Promise<void> => {
    const receiverRocketChatUserId =
        inBoundNotification.receiverRocketChatUserId;

    const resourceString = inBoundNotification.resourceString;
    const getMessageResponse = await getMessageWithResourceStringAsync(
        http,
        resourceString,
        userAccessToken,
    );

    const messageIdMapping = await MessageMapping.findByTeamsMessageId(
        read,
        getMessageResponse.messageId,
    );
    if (!messageIdMapping) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }

    if (
        await PreventRegistry.capture(
            persis,
            `PreventPostMessageUpdateHook/${messageIdMapping.rocketChatMessageId}`,
        )
    ) {
        return;
    }
    const fromUserTeamsId = getMessageResponse.fromUserTeamsId;
    if (!fromUserTeamsId) {
        // If there's not a sender, stop processing
        return;
    }

    const message = await read
        .getMessageReader()
        .getById(messageIdMapping.rocketChatMessageId);
    if (!message) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }
    const appUser = await read.getUserReader().getAppUser();
    let botFallback = false;
    let displayName = `Teams User (${fromUserTeamsId})`;

    if (appUser) {
        const appUsermapping = await UserMapping.findByRCUserId(
            read,
            appUser.id,
        );
        if (
            appUsermapping?.teamsUserId !== fromUserTeamsId &&
            message.sender.id === appUser.id
        ) {
            botFallback = true;
            const accessToken = await getUserAccessTokenAsync({
                read,
                persistence: persis,
                rocketChatUserId: appUser.id,
                app,
                http,
            });
            if (accessToken) {
                const senderProfile = await getTeamsUserProfileByIdAsync(
                    http,
                    accessToken,
                    fromUserTeamsId,
                );
                displayName = senderProfile?.displayName ?? displayName;
            }
        }
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
    let messageBuilder = await updator.message(
        messageIdMapping.rocketChatMessageId,
        sender,
    );

    messageBuilder = messageBuilder
        .setText(
            botFallback
                ? formatTeamsSenderInfo(updatedMessage.text, displayName)
                : updatedMessage.text,
        )
        .setEditor(sender);
    await PreventRegistry.set(
        persis,
        `PreventPostMessageUpdateHook/${messageIdMapping.rocketChatMessageId}`,
    );
    await updator.finish(messageBuilder);
};
