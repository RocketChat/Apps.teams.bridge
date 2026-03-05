import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { DefaultTeamName } from "./Const";
import {
    mapTeamsMessageToRocketChatMessage,
    sendRocketChatMessageInRoomAsync,
} from "./MessageHelper";
import {
    getChatThreadWithMembersAsync,
    getMessageWithResourceStringAsync,
    MessageType,
    ThreadType,
} from "./MicrosoftGraphApi";
import { MessageMapping, Room, TeamsUserProfile, UploadMapping, UserMapping } from "./PersistHelper";
import type { UserModel } from "./PersistHelper";
import type { TeamsBridgeApp } from "../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "./AuthHelper";
import { PreventRegistry } from "./PreventRegistry";

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
                app.getID(),
            );
            break;

        case NotificationChangeType.Updated:
            await handleInboundMessageUpdatedAsync(
                userAccessToken,
                inBoundNotification,
                read,
                modify,
                http,
                persistence
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

const handleInboundMessageCreatedAsync = async (
    userAccessToken: string,
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence,
    appId: string,
): Promise<void> => {
    const receiverRocketChatUserId = inBoundNotification.receiverRocketChatUserId;
    const resourceString = inBoundNotification.resourceString;
    const getMessageResponse = await getMessageWithResourceStringAsync(
        http,
        resourceString,
        userAccessToken
    );

    if (getMessageResponse.messageType) {
        const storedMessageMap = await MessageMapping.findByTeamsMessageId(read, getMessageResponse.messageId);

        if (storedMessageMap?.rocketChatMessageId) {
            // IMPORTANT!!!!!
            // An echo message. Should skip. Else this will create a loop.
            return;
        }

        let roomRecord = await Room.findByTeamsThreadId(
            read,
            getMessageResponse.threadId
        );
        if (!roomRecord) {
            if (getMessageResponse.messageType !== MessageType.Message) {
                // Only create room for real message
                return;
            }

            // Handle thread created in Teams scenario
            // Get thread and members info
            const threadInfo = await getChatThreadWithMembersAsync(
                http,
                getMessageResponse.threadId,
                userAccessToken
            );

            // Build a room with thread info
            const userReader = read.getUserReader();
            const notificationReceiverUser = await userReader.getById(
                receiverRocketChatUserId
            );

            let topic = DefaultTeamName;

            const creator = modify.getCreator();
            const roomBuilder = creator.startRoom();
            roomBuilder.setCreator(notificationReceiverUser);
            if (threadInfo.type) {
                if (threadInfo.type === ThreadType.OneOnOne) {
                    roomBuilder
                        .setType(RoomType.DIRECT_MESSAGE)
                        .setSlugifiedName(`dm_${notificationReceiverUser.id}`);
                } else if (threadInfo.type === ThreadType.Group) {
                    roomBuilder
                        .setType(RoomType.PRIVATE_GROUP)
                        .setDisplayName(topic)
                        .setSlugifiedName(topic);
                } else {
                    throw new Error(
                        `Unsupported thread type ${threadInfo.type} found for Teams thread ${threadInfo.threadId}`
                    );
                }

                const teamsMemberIds = threadInfo.memberIds;
                if (!teamsMemberIds || teamsMemberIds.length == 0) {
                    throw new Error(
                        `No members found for Teams thread ${threadInfo.threadId}`
                    );
                }

                // Add thread members to the room
                for (const teamsMemberId of teamsMemberIds) {
                    const rocketChatUser = await UserMapping.findByTeamsUserId(
                        read,
                        teamsMemberId
                    );
                    if (rocketChatUser) {
                        const user = await userReader.getById(
                            rocketChatUser.rocketChatUserId
                        );
                        roomBuilder.addMemberToBeAddedByUsername(user.username);
                    } else {
                        // Under single-bot arch there are no dummy users. Teams-only members
                        // who have no RC registration are not added to the RC room.
                        console.log(
                            `No RC user found for Teams member ${teamsMemberId}, skipping room membership.`
                        );
                    }
                }
            } else {
                throw new Error(
                    `No thread type found for Teams thread ${threadInfo.threadId}`
                );
            }

            const roomId = await creator.finish(roomBuilder);
            console.log(`Room ${roomId} created for incoming message!`);

            // Persist room record
            await Room.persist(
                persis,
                roomId,
                threadInfo.threadId
            );

            roomRecord = await Room.findByTeamsThreadId(
                read,
                getMessageResponse.threadId
            );
            if (!roomRecord) {
                throw new Error(
                    `Create room failed for Teams thread ${getMessageResponse.threadId}`
                );
            }
        }

        const room = await read
            .getRoomReader()
            .getById(roomRecord.rocketChatRoomId);
        if (!room) {
            return;
        }

        // Only handle notification received by the app bot to avoid duplication
        const appUser = await read.getUserReader().getAppUser(appId);
        if (receiverRocketChatUserId !== appUser?.id) {
            console.log("Skip notification for non-app user");
            return;
        }

        if (getMessageResponse.messageType === MessageType.Message) {
            const fromUserTeamsId = getMessageResponse.fromUserTeamsId;
            if (!fromUserTeamsId) {
                // If there's no sender, stop processing
                console.error("No sender for message");
                return;
            }

            const fromUserRocketChatUser = await UserMapping.findByTeamsUserId(
                read,
                fromUserTeamsId
            );

            const senderUser = await getSenderUser({
                roomRecord,
                fromUserRocketChatUser,
                read,
                appId,
                fromUserTeamsId,
            });

            if (!senderUser) {
                throw new Error('No user found to send the message');
            }

            // When the sender has no RC registration the app bot relays the message.
            // Prefix the message text with the Teams sender's display name so RC
            // users can see who originally sent it.
            const usesBotFallback = !fromUserRocketChatUser;

            const message = await mapTeamsMessageToRocketChatMessage({
                getMessageResponse,
                accessToken: userAccessToken,
                room,
                sender: senderUser,
                http,
                modify,
                read,
                uploadFiles: true,
            });

            if (usesBotFallback && message.text !== "") {
                const senderProfile = await TeamsUserProfile.findByTeamsUserId(read, fromUserTeamsId);
                const displayName = senderProfile?.displayName ?? fromUserTeamsId;
                message.text = `**${displayName}:** ${message.text}`;
            }

            if (message.text === "") {
                // File message, no text content
                await Promise.all(
                    message.uploadIds.map((uploadIdMap) => {
                        return UploadMapping.persist({
                            persistence: persis,
                            rocketchatUploadId: uploadIdMap.rocketChat,
                            teamsAttachmentId: uploadIdMap.teams,
                            teamsMessageId: getMessageResponse.messageId,
                            teamsThreadId: getMessageResponse.threadId,
                        });
                    }),
                );
                return;
            }

            const rocketChatMessageId = await sendRocketChatMessageInRoomAsync(
                message.text,
                senderUser,
                room,
                modify
            );

            await Promise.all([
                MessageMapping.persist({
                    persistence: persis,
                    rocketChatMessageId,
                    teamsMessageId: getMessageResponse.messageId,
                    teamsThreadId: getMessageResponse.threadId,
                }),
                ...message.uploadIds.map((uploadIdMap) => {
                    return UploadMapping.persist({
                        persistence: persis,
                        rocketchatUploadId: uploadIdMap.rocketChat,
                        teamsAttachmentId: uploadIdMap.teams,
                        teamsMessageId: getMessageResponse.messageId,
                        teamsThreadId: getMessageResponse.threadId,
                    });
                }),
            ]);
        } else if (
            getMessageResponse.messageType === MessageType.SystemAddMembers
        ) {
            const memberToAddTeamsIds = getMessageResponse.memberIds;
            if (!memberToAddTeamsIds || memberToAddTeamsIds.length === 0) {
                console.error("Empty members Id list for add members.");
                return;
            }

            for (const memberToAddTeamsId of memberToAddTeamsIds) {
                let userToAdd: IUser | undefined = undefined;

                // First, try find whether there's a real Rocket.Chat user for this Teams user to add
                const rocketChatUser = await UserMapping.findByTeamsUserId(
                    read,
                    memberToAddTeamsId
                );
                if (rocketChatUser) {
                    userToAdd = await read
                        .getUserReader()
                        .getById(rocketChatUser.rocketChatUserId);
                } else {
                    // Under single-bot arch there are no dummy users. Teams members without
                    // a registered RC account are not added to the RC room.
                    console.log(
                        `No RC user found for Teams member ${memberToAddTeamsId}, skipping room membership.`
                    );
                    continue;
                }

                const updater = modify.getUpdater();
                const roomBuilder = await updater.room(room.id, room.creator);

                if (!userToAdd) {
                    console.error("Could not add Teams bot user to room!");
                    console.error(
                        `Dummy user with Teams ID ${memberToAddTeamsId} not found after try sync all Teams bot users!`
                    );
                    continue;
                }

                roomBuilder.addMemberToBeAddedByUsername(userToAdd.username);
                await updater.finish(roomBuilder);
            }
        } else {
            console.log("Unsupported message type.");
        }
    } else {
        console.log("Unsupported message type.");
    }
};

const getSenderUser = async ({
        roomRecord,
        fromUserRocketChatUser,
        read,
        appId,
        fromUserTeamsId,
    }: {
        roomRecord: any,
        fromUserRocketChatUser: UserModel | null,
        read: IRead,
        appId: string,
        fromUserTeamsId: string,
}) => {
    if (fromUserRocketChatUser) {
        const roomMembers = await read.getRoomReader().getMembers(roomRecord.rocketChatRoomId);
        if (roomMembers && roomMembers.find((user) => user.id === fromUserRocketChatUser.rocketChatUserId)) {
            return read.getUserReader().getById(fromUserRocketChatUser.rocketChatUserId);
        }
    }

    // Under single-bot arch there are no dummy users. Fall back to the app bot
    // so the message is still relayed to RC under the bridge bot identity.
    console.log(
        `No RC user found for Teams sender ${fromUserTeamsId}, falling back to app bot.`
    );
    return read.getUserReader().getAppUser(appId);
}

const handleInboundMessageUpdatedAsync = async (
    userAccessToken: string,
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence
): Promise<void> => {
    const receiverRocketChatUserId =
        inBoundNotification.receiverRocketChatUserId;

    const resourceString = inBoundNotification.resourceString;
    const getMessageResponse = await getMessageWithResourceStringAsync(
        http,
        resourceString,
        userAccessToken
    );

    const messageIdMapping =
        await MessageMapping.findByTeamsMessageId(
            read,
            getMessageResponse.messageId
        );
    if (!messageIdMapping) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }

    if (
        await PreventRegistry.capture(
            persis,
            `PreventPostMessageUpdateHook/${messageIdMapping.rocketChatMessageId}`
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
    });

    const updator = modify.getUpdater();
    let messageBuilder = await updator.message(
        messageIdMapping.rocketChatMessageId,
        sender
    );
    messageBuilder = messageBuilder
        .setText(updatedMessage.text)
        .setEditor(sender)
    await updator.finish(messageBuilder);
};

const handleInboundMessageDeletedAsync = async (
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence
): Promise<void> => {
    const resourceString = inBoundNotification.resourceId;

    const messageIdMapping =
        await MessageMapping.findByTeamsMessageId(
            read,
            resourceString
        );

    if (!messageIdMapping) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }

    if (
        await PreventRegistry.capture(
            persis,
            `PreventPostMessageDeleteHook/${messageIdMapping.rocketChatMessageId}`
        )
    ) {
        // Prevent duplicate processing
        return;
    }

    const message = await read
        .getMessageReader()
        .getById(messageIdMapping.rocketChatMessageId);
    if (!message) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }

    const sender: IUser = message.sender;

    const updator = modify.getUpdater();
    let messageBuilder = await updator.message(
        messageIdMapping.rocketChatMessageId,
        sender
    );

    await PreventRegistry.set(
        persis,
        `PreventPostMessageUpdateHook/${messageIdMapping.rocketChatMessageId}`
    );
    messageBuilder = messageBuilder
        .setText("~This message has been deleted.~")
        .setEditor(sender);
    await updator.finish(messageBuilder);
};
