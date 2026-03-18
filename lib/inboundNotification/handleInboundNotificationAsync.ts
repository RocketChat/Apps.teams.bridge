import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import type { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "./../AuthHelper";
import { handleInboundMessageCreatedAsync } from "./handleInboundMessageCreatedAsync";
import { handleInboundMessageUpdatedAsync } from "./handleInboundMessageUpdatedAsync";
import { handleInboundMessageDeletedAsync } from "./handleInboundMessageDeletedAsync";

export enum NotificationChangeType {
    Created = "created",
    Updated = "updated",
    Deleted = "deleted",
}

export enum NotificationResourceType {
    ChatMessage = "chatMessage",
}

export interface InBoundNotification {
    receiverRocketChatUserId: string;
    subscriptionId: string;
    changeType: NotificationChangeType;
    resourceId: string;
    resourceString: string;
    resourceType: NotificationResourceType;
}

export const handleInboundNotificationAsync = async (options: {
    inBoundNotification: InBoundNotification;
    read: IRead;
    modify: IModify;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const { app, http, inBoundNotification, modify, persistence, read } = options;
    const receiverRocketChatUserId =
        inBoundNotification.receiverRocketChatUserId;
    if (!receiverRocketChatUserId) {
        // If there's not a receiver, stop processing
        return;
    }

    const appUser = await read.getUserReader().getAppUser();
    if (receiverRocketChatUserId !== appUser?.id) {
        console.log("Skip notification for non-app user");
        return;
    }

    const userAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: receiverRocketChatUserId,
        app,
        http,
    });

    if (!userAccessToken) {
        // If receiver's access token does not exist in persist or expired, stop processing
        console.error(
            `Receiver user ${receiverRocketChatUserId} access token does not exist in persist or expired`
        );
        return;
    }

    switch (inBoundNotification.changeType) {
        case NotificationChangeType.Created:
            await handleInboundMessageCreatedAsync(
                userAccessToken,
                inBoundNotification,
                read,
                modify,
                http,
                persistence,
                app,
            );
            break;

        case NotificationChangeType.Updated:
            await handleInboundMessageUpdatedAsync(
                userAccessToken,
                inBoundNotification,
                read,
                modify,
                http,
                persistence,
                app,
            );
            break;

        case NotificationChangeType.Deleted:
            await handleInboundMessageDeletedAsync(
                inBoundNotification,
                read,
                modify,
                http,
                persistence
            );
            break;

        default:
            console.error(`Unsupported notification change type`);
            return;
    }
};
