import {
    IHttp,
    IModify,
    IPersistence,
    IRead
} from "@rocket.chat/apps-engine/definition/accessors";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { syncAllTeamsBotUsersAsync } from "./AppUserHelper";
import { DefaultTeamName, DefaultThreadName } from "./Const";
import { mapTeamsMessageToRocketChatMessage, sendRocketChatMessageInRoomAsync, sendRocketChatOneOnOneMessageAsync } from "./MessageHelper";
import { getChatThreadWithMembersAsync, getMessageWithResourceStringAsync, MessageContentType, MessageType, ThreadType } from "./MicrosoftGraphApi";
import {
    persistMessageIdMappingAsync,
    persistRoomAsync,
    retrieveDummyUserByTeamsUserIdAsync,
    retrieveMessageIdMappingByTeamsMessageIdAsync,
    retrieveRoomByTeamsThreadIdAsync,
    retrieveUserAccessTokenAsync,
    retrieveUserByTeamsUserIdAsync
} from "./PersistHelper";

export enum NotificationChangeType {
    Created = 'created',
    Updated = 'updated',
    Deleted = 'deleted',
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
    const receiverRocketChatUserId = inBoundNotification.receiverRocketChatUserId;
    if (!receiverRocketChatUserId) {
        // If there's not a receiver, stop processing
        return;
    }

    const userAccessToken = await retrieveUserAccessTokenAsync(read, receiverRocketChatUserId);
    if (!userAccessToken) {
        // If receiver's access token does not exist in persist or expired, stop processing
        console.error(`Receiver user ${receiverRocketChatUserId} access token does not exist in persist or expired`);
        return;
    }

    switch (inBoundNotification.changeType) {
        case NotificationChangeType.Created:
            await handleInboundMessageCreatedAsync(userAccessToken, inBoundNotification, read, modify, http, persis);
            break;
        
        case NotificationChangeType.Updated:
            await handleInboundMessageUpdatedAsync(userAccessToken, inBoundNotification, read, modify, http, persis);
            break;

        case NotificationChangeType.Deleted:
            await handleInboundMessageDeletedAsync(inBoundNotification, read, modify, http, persis);
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
    persis: IPersistence) : Promise<void> => {
    const receiverRocketChatUserId = inBoundNotification.receiverRocketChatUserId;
    const resourceString = inBoundNotification.resourceString;
    const getMessageResponse = await getMessageWithResourceStringAsync(http, resourceString, userAccessToken);

    if (getMessageResponse.messageType) {
                
        let roomRecord = await retrieveRoomByTeamsThreadIdAsync(read, getMessageResponse.threadId);
        if (!roomRecord) {
            if (getMessageResponse.messageType !== MessageType.Message) {
                // Only create room for real message
                return;
            }

            console.log("Create new room for incoming message!");
            // Handle thread created in Teams scenario
            // Get thread and members info
            const threadInfo = await getChatThreadWithMembersAsync(http, getMessageResponse.threadId, userAccessToken);

            // Build a room with thread info
            const userReader = read.getUserReader();
            const notificationReceiverUser = await userReader.getById(receiverRocketChatUserId);

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
                    throw new Error(`Unsupported thread type ${threadInfo.type} found for Teams thread ${threadInfo.threadId}`);
                }

                const teamsMemberIds = threadInfo.memberIds;
                if (!teamsMemberIds || teamsMemberIds.length == 0) {
                    throw new Error(`No members found for Teams thread ${threadInfo.threadId}`);
                }

                // Add thread members to the room
                for (const teamsMemberId of teamsMemberIds) {
                    const rocketChatUser = await retrieveUserByTeamsUserIdAsync(read, teamsMemberId);
                    if (rocketChatUser) {
                        const user = await userReader.getById(rocketChatUser.rocketChatUserId);
                        roomBuilder.addMemberToBeAddedByUsername(user.username);
                    } else {
                        const dummyUser = await retrieveDummyUserByTeamsUserIdAsync(read, teamsMemberId);
                        if (!dummyUser) {
                            console.error(`No dummy user found for Teams user ${teamsMemberId}, skip.`)
                            continue;
                        }

                        const user = await userReader.getById(dummyUser.rocketChatUserId);
                        roomBuilder.addMemberToBeAddedByUsername(user.username);
                    }
                }
            } else {
                throw new Error(`No thread type found for Teams thread ${threadInfo.threadId}`);
            }

            const roomId = await creator.finish(roomBuilder);
            console.log(`Room ${roomId} created for incoming message!`);

            // Set notification receiver as bridge user and persist room record
            await persistRoomAsync(persis, roomId, threadInfo.threadId, receiverRocketChatUserId);

            roomRecord = await retrieveRoomByTeamsThreadIdAsync(read, getMessageResponse.threadId);
            if (!roomRecord) {
                throw new Error(`Create room failed for Teams thread ${getMessageResponse.threadId}`);
            }
        }

        const room = await read.getRoomReader().getById(roomRecord.rocketChatRoomId);
        if (!room) {
            return;
        }

        // Only handle notification received by the bridge user to avoid duplication
        if (!roomRecord.bridgeUserRocketChatUserId || roomRecord.bridgeUserRocketChatUserId !== receiverRocketChatUserId) {
            console.log("Skip notification for non-bridge user");
            return;
        }

        if (getMessageResponse.messageType === MessageType.Message) {
            const fromUserTeamsId = getMessageResponse.fromUserTeamsId;
            if (!fromUserTeamsId) {
                // If there's not a sender, stop processing
                console.error("No sender for message");
                return;
            }
    
            const fromUserRocketChatUser = await retrieveUserByTeamsUserIdAsync(read, fromUserTeamsId);
            if (fromUserRocketChatUser) {
                // If this message is sent from a Teams user that his/her corresponding real Rocket Chat user is in current room, stop processing
                const roomMembers = await read.getRoomReader().getMembers(roomRecord.rocketChatRoomId);
                if (roomMembers && roomMembers.find(user => user.id === fromUserRocketChatUser.rocketChatUserId)) {
                    console.log("Message not from dummy user!");
                    return;
                }
            }

            let fromDummyUser = await retrieveDummyUserByTeamsUserIdAsync(read, fromUserTeamsId);
            if (!fromDummyUser) {
                // There could be dummy user out of sync issue.
                // If the dummy user has not been created for a recently added Teams user, we need to create dummy user on demand.
                // Sync all Teams bot user
                await syncAllTeamsBotUsersAsync(http, read, modify, persis);
                fromDummyUser = await retrieveDummyUserByTeamsUserIdAsync(read, fromUserTeamsId);
                if (!fromDummyUser) {
                    throw new Error(`Dummy user with Teams ID ${fromUserTeamsId} not found after try sync all Teams bot users!`);
                }
            }
    
            const sender: IUser = await read.getUserReader().getById(fromDummyUser.rocketChatUserId);
    
            const messageText = mapTeamsMessageToRocketChatMessage(
                getMessageResponse,
                userAccessToken,
                room,
                sender,
                http,
                modify);
    
            if (messageText === '') {
                // File message, no text content
                return;
            }
            
            const rocketChatMessageId = await sendRocketChatMessageInRoomAsync(messageText, sender, room, modify);
            await persistMessageIdMappingAsync(persis, rocketChatMessageId, getMessageResponse.messageId, getMessageResponse.threadId);

        } else if (getMessageResponse.messageType === MessageType.SystemAddMembers) {
            const memberToAddTeamsIds = getMessageResponse.memberIds;
            if (!memberToAddTeamsIds || memberToAddTeamsIds.length === 0) {
                console.error("Empty members Id list for add members.");
                return;
            }

            for (const memberToAddTeamsId of memberToAddTeamsIds) {
                let userToAdd : IUser | undefined = undefined;

                // First, try find whether there's a real Rocket.Chat user for this Teams user to add
                const rocketChatUser = await retrieveUserByTeamsUserIdAsync(read, memberToAddTeamsId);
                if (rocketChatUser) {
                    userToAdd = await read.getUserReader().getById(rocketChatUser.rocketChatUserId);
                } else {
                    // If there's not, try find the Teams bot user and add to the Rocket.Chat room
                    let dummyUser = await retrieveDummyUserByTeamsUserIdAsync(read, memberToAddTeamsId);
                    if (!dummyUser) {
                        // There could be dummy user out of sync issue.
                        // If the dummy user has not been created for a recently added Teams user, we need to create dummy user on demand.
                        // Sync all Teams bot user
                        await syncAllTeamsBotUsersAsync(http, read, modify, persis);
                        dummyUser = await retrieveDummyUserByTeamsUserIdAsync(read, memberToAddTeamsId);
                        if (!dummyUser) {
                            console.error('Could not add Teams bot user to room!');
                            console.error(`Dummy user with Teams ID ${memberToAddTeamsId} not found after try sync all Teams bot users!`);
                            continue;
                        }
                    }

                    userToAdd = await read.getUserReader().getById(dummyUser.rocketChatUserId);
                }

                const updater = modify.getUpdater();
                const roomBuilder = await updater.room(room.id, room.creator);
        
                if (!userToAdd) {
                    console.error('Could not add Teams bot user to room!');
                    console.error(`Dummy user with Teams ID ${memberToAddTeamsId} not found after try sync all Teams bot users!`);
                    continue;
                }
        
                roomBuilder.addMemberToBeAddedByUsername(userToAdd.username);
                await updater.finish(roomBuilder);
            }
        } else {
            console.log('Unsupported message type.');
        }
    } else {
        console.log('Unsupported message type.');
    }
};

const handleInboundMessageUpdatedAsync = async (
    userAccessToken: string,
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence) : Promise<void> => {
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
    if (!message) {
        // If there's not an existing rocket chat message, stop processing
        return;
    }

    const sender: IUser = message.sender;
    const updatedMessageText = mapTeamsMessageToRocketChatMessage(
        getMessageResponse,
        userAccessToken,
        message.room,
        sender,
        http,
        modify);
        
    const updator = modify.getUpdater();
    let messageBuilder = await updator.message(messageIdMapping.rocketChatMessageId, sender);
    messageBuilder = messageBuilder
        .setText(updatedMessageText)
        .setEditor(sender);
    await updator.finish(messageBuilder);
};

const handleInboundMessageDeletedAsync = async (
    inBoundNotification: InBoundNotification,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence) : Promise<void> => {
    const resourceString = inBoundNotification.resourceId;

    const messageIdMapping = await retrieveMessageIdMappingByTeamsMessageIdAsync(read, resourceString);
    if (!messageIdMapping) {
        // If there's not an existing rocket chat message, stop processing
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
    messageBuilder = messageBuilder
        .setText('~This message has been deleted.~')
        .setEditor(sender);
    await updator.finish(messageBuilder);
};
