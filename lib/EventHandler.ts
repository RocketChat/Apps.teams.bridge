import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { UserNotAllowedException } from "@rocket.chat/apps-engine/definition/exceptions";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import {
    IRoom,
    IRoomUserLeaveContext,
    RoomType,
} from "@rocket.chat/apps-engine/definition/rooms";
import { IFileUploadContext } from "@rocket.chat/apps-engine/definition/uploads";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import { findAllDummyUsersInRocketChatUserListAsync } from "./AppUserHelper";
import {
    AddUserLoginRequiredHintMessageText,
    AuthenticationEndpointPath,
    BridgeUserNotificationMessageText,
    DefaultThreadName,
    LoggedInBridgeUserRequiredHintMessageText,
    LoginRequiredHintMessageText,
    SubscriberEndpointPath,
    UnsupportedScenarioHintMessageText,
} from "./Const";
import {
    generateHintMessageWithTeamsLoginButton,
    mapRocketChatMessageToTeamsMessage,
    notifyRocketChatUserAsync,
    notifyRocketChatUserInRoomAsync,
} from "./MessageHelper";
import {
    addMemberToChatThreadAsync,
    createChatThreadAsync,
    createOneOnOneChatThreadAsync,
    deleteTextMessageInChatThreadAsync,
    listMembersInChatThreadAsync,
    listSubscriptionsAsync,
    removeMemberFromChatThreadAsync,
    renewSubscriptionAsync,
    renewUserAccessTokenAsync,
    sendFileMessageToChatThreadAsync,
    sendTextMessageToChatThreadAsync,
    shareOneDriveFileAsync,
    subscribeToAllMessagesForOneUserAsync,
    updateTextMessageInChatThreadAsync,
    uploadFileToOneDriveAsync,
} from "./MicrosoftGraphApi";
import {
    checkDummyUserByRocketChatUserIdAsync,
    persistMessageIdMappingAsync,
    persistOneDriveFileAsync,
    persistRoomAsync,
    persistUserAccessTokenAsync,
    retrieveAllUserRegistrationsAsync,
    retrieveDummyUserByRocketChatUserIdAsync,
    retrieveDummyUserByTeamsUserIdAsync,
    retrieveLoginMessageSentStatus,
    retrieveMessageIdMappingByRocketChatMessageIdAsync,
    retrieveOneDriveFileAsync,
    retrieveRoomByRocketChatRoomIdAsync,
    retrieveUserAccessTokenAsync,
    retrieveUserByRocketChatUserIdAsync,
    retrieveUserRefreshTokenAsync,
    saveLoginMessageSentStatus,
    UserModel,
} from "./PersistHelper";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "./UrlHelper";

export const handlePreMessageSentPreventAsync = async (
    message: IMessage,
    read: IRead,
    persistence: IPersistence,
    app: TeamsBridgeApp
): Promise<boolean> => {
    const wasSent = await retrieveLoginMessageSentStatus({
        read,
        rocketChatUserId: message.sender.id,
    });

    if (wasSent) {
        return false;
    }

    const appUser = (await read.getUserReader().getByUsername('microsoftteamsbridge.bot')) as IUser;
    const notifier = read.getNotifier();

    if (message.threadId) {
        const isTeamsMessageThread = await isTeamsMessageAsync(
            message.threadId,
            read
        );
        if (isTeamsMessageThread) {
            // There's no thread message concept in Teams
            // Thread message is not a supported scenario for Teams interop

            await notifyRocketChatUserInRoomAsync(
                UnsupportedScenarioHintMessageText("Thread Message"),
                appUser,
                message.sender,
                message.room,
                notifier
            );

            return true;
        }
    }

    const roomType = message.room.type;
    if (
        roomType === RoomType.PRIVATE_GROUP ||
        roomType === RoomType.DIRECT_MESSAGE
    ) {
        // If room type is PRIVATE_GROUP or DIRECT_MESSAGE, check if there's any dummy user in the room
        const members = await read.getRoomReader().getMembers(message.room.id);

        const dummyUsers = await findAllDummyUsersInRocketChatUserListAsync(
            read,
            members
        );

        if (dummyUsers && dummyUsers.length > 0) {
            // If there are dummy users in the room, check whether there's at least one teams-logged in user in this room

            // Find whether there's an existing room record
            let roomRecord = await retrieveRoomByRocketChatRoomIdAsync(
                read,
                message.room.id
            );

            if (roomRecord) {
                // If there's an existing room record, check whether it has a bridge user
                if (roomRecord.bridgeUserRocketChatUserId) {
                    // If this room already has assigned a bridge user, check the bridge user login status

                    const accessToken = await retrieveUserAccessTokenAsync(
                        read,
                        persistence,
                        roomRecord.bridgeUserRocketChatUserId
                    );
                    if (!accessToken) {
                        // If the existing bridge user is logged out, clean the bridge user
                        roomRecord.bridgeUserRocketChatUserId = undefined;
                    }
                }
            } else {
                // Create a new room record if there's not an existing one
                roomRecord = {
                    rocketChatRoomId: message.room.id,
                };
            }

            // Try to find a logged in user and assign to the room as the bridge user
            if (!roomRecord.bridgeUserRocketChatUserId) {
                const loggedInUser = await findOneTeamsLoggedInUsersAsync(
                    read,
                    persistence,
                    members
                );
                const isOneOnOneDirectMessage =
                    roomType === RoomType.DIRECT_MESSAGE &&
                    members.length === 2;
                if (loggedInUser) {
                    // Assign the room a bridge user
                    roomRecord.bridgeUserRocketChatUserId =
                        loggedInUser.rocketChatUserId;

                    // For 1:1 dm chat, no further action required
                    if (!isOneOnOneDirectMessage) {
                        // For other type of chat room
                        // Notify the bridge user that he has became the bridge of this room
                        // All messages sent by unlogged in user will be delivered to Microsoft Teams by him
                        const bridgeUser = await read
                            .getUserReader()
                            .getById(loggedInUser.rocketChatUserId);
                        await notifyRocketChatUserInRoomAsync(
                            BridgeUserNotificationMessageText,
                            appUser,
                            bridgeUser,
                            message.room,
                            notifier
                        );

                        // TODO: send a message to Microsoft Teams to let the user there know the bridge user represents some other users
                    }
                } else {
                    // If there's no logged in user in the room, prevent the message
                    if (isOneOnOneDirectMessage) {
                        // For 1:1 chat, notify the sender to login

                        await notifyNotLoggedInUserAsync(
                            read,
                            message.sender,
                            message.room,
                            app,
                            LoginRequiredHintMessageText
                        );
                        await saveLoginMessageSentStatus({
                            persistence,
                            rocketChatUserId: message.sender.id,
                            wasSent: true,
                        });
                    } else {
                        // For other type of chat room
                        // Notify the message sender there's no available bridge user

                        await notifyNotLoggedInUserAsync(
                            read,
                            message.sender,
                            message.room,
                            app,
                            LoggedInBridgeUserRequiredHintMessageText
                        );
                        await saveLoginMessageSentStatus({
                            persistence,
                            rocketChatUserId: message.sender.id,
                            wasSent: true,
                        });
                    }
                }
            }

            // Persist the room record
            await persistRoomAsync(
                persistence,
                roomRecord.rocketChatRoomId,
                roomRecord.teamsThreadId,
                roomRecord.bridgeUserRocketChatUserId
            );
        }
    }

    return false;
};

export const handlePostMessageSentAsync = async (
    message: IMessage,
    read: IRead,
    http: IHttp,
    persistence: IPersistence
): Promise<void> => {
    const isSenderDummyUser = await checkDummyUserByRocketChatUserIdAsync(
        read,
        message.sender.id
    );
    if (isSenderDummyUser) {
        console.log("Message sender is a dummy user, stop processing.");
        return;
    }

    const roomId = message.room.id;
    const members = await read.getRoomReader().getMembers(roomId);

    const dummyUsers = await findAllDummyUsersInRocketChatUserListAsync(
        read,
        members
    );
    if (dummyUsers && dummyUsers.length > 0) {
        // If there's any dummy user in the room, this is a Teams interop chat room
        // Sanity check has been done in PreMessageSentPrevent for Teams interop scenarios

        // There should be a room record in persist with a bridge user assigned
        const roomRecord = await retrieveRoomByRocketChatRoomIdAsync(
            read,
            roomId
        );
        if (!roomRecord) {
            throw new Error("No room record find for Teams interop room!");
        }

        if (!roomRecord.bridgeUserRocketChatUserId) {
            throw new Error("No bridge user assigned to Teams interop room!");
        }

        const bridgeUser = await retrieveUserByRocketChatUserIdAsync(
            read,
            roomRecord.bridgeUserRocketChatUserId
        );
        let userAccessToken = await retrieveUserAccessTokenAsync(
            read,
            persistence,
            roomRecord.bridgeUserRocketChatUserId
        );
        if (!userAccessToken || !bridgeUser) {
            await persistRoomAsync(
                persistence,
                roomRecord.rocketChatRoomId,
                roomRecord.teamsThreadId,
                undefined
            );
            throw new Error("Invalid bridge user!");
        }

        if (!roomRecord.teamsThreadId) {
            // Not yet a thread exist in Teams side, create one & persist in room record
            if (
                message.room.type === RoomType.DIRECT_MESSAGE &&
                members.length === 2
            ) {
                // If 1:1 DM, create 1:1 Teams chat thread
                const response = await createOneOnOneChatThreadAsync(
                    http,
                    bridgeUser.teamsUserId,
                    dummyUsers[0].teamsUserId,
                    userAccessToken
                );
                roomRecord.teamsThreadId = response.threadId;
            } else {
                // If other room type, create Teams group chat thread
                const teamsIds: string[] = [];
                for (const member of members) {
                    const user = await retrieveUserByRocketChatUserIdAsync(
                        read,
                        member.id
                    );
                    if (user) {
                        teamsIds.push(user.teamsUserId);
                    }
                }

                for (const dummyUser of dummyUsers) {
                    teamsIds.push(dummyUser.teamsUserId);
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

            await persistRoomAsync(
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

        const senderUserAccessToken = await retrieveUserAccessTokenAsync(
            read,
            persistence,
            message.sender.id
        );
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

            const oneDriveFile = await retrieveOneDriveFileAsync(
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
            // Mapping message content format
            messageText = mapRocketChatMessageToTeamsMessage(
                messageText,
                originalSenderName
            );

            // Send the message to the chat thread
            const response = await sendTextMessageToChatThreadAsync(
                http,
                messageText,
                roomRecord.teamsThreadId,
                userAccessToken
            );

            teamsMessageId = response.messageId;
            rocketChatMessageId = message.id as string;
        }

        await persistMessageIdMappingAsync(
            persistence,
            rocketChatMessageId,
            teamsMessageId,
            roomRecord.teamsThreadId
        );
    }
};

export const handlePreMessageOperationPreventAsync = async (
    message: IMessage,
    read: IRead,
    persistence: IPersistence
): Promise<boolean> => {
    const isTeamsMessage = await isTeamsMessageAsync(message.id, read);
    if (!isTeamsMessage) {
        return false;
    }

    // If the user that operate the Teams message has not logged in to Teams
    // Send a notification to let the sender know he need to logged in to Teams to apply the operation
    const senderId = message.sender.id;
    const dummyUser = await retrieveDummyUserByRocketChatUserIdAsync(
        read,
        senderId
    );
    if (dummyUser) {
        return false;
    }

    const senderUserAccessToken = await retrieveUserAccessTokenAsync(
        read,
        persistence,
        senderId
    );
    if (!senderUserAccessToken) {
        return true;
    }

    return false;
};

export const handlePostMessageUpdatedAsync = async (
    message: IMessage,
    read: IRead,
    persistence: IPersistence,
    http: IHttp
): Promise<void> => {
    if (!message || !message.id || !message.text) {
        return;
    }

    const messageIdMapping =
        await retrieveMessageIdMappingByRocketChatMessageIdAsync(
            read,
            message.id
        );
    if (!messageIdMapping) {
        return;
    }

    const senderId = message.sender.id;
    const senderUserAccessToken = await retrieveUserAccessTokenAsync(
        read,
        persistence,
        senderId
    );
    if (!senderUserAccessToken) {
        return;
    }

    await updateTextMessageInChatThreadAsync(
        http,
        message.text,
        messageIdMapping.teamsMessageId,
        messageIdMapping.teamsThreadId,
        senderUserAccessToken
    );
};

export const handlePostMessageDeletedAsync = async (
    message: IMessage,
    read: IRead,
    persistence: IPersistence,
    http: IHttp
): Promise<void> => {
    if (!message || !message.id || !message.text) {
        return;
    }

    const messageIdMapping =
        await retrieveMessageIdMappingByRocketChatMessageIdAsync(
            read,
            message.id
        );
    if (!messageIdMapping) {
        return;
    }

    const senderId = message.sender.id;
    const senderUserAccessToken = await retrieveUserAccessTokenAsync(
        read,
        persistence,
        senderId
    );
    if (!senderUserAccessToken) {
        return;
    }

    const senderUser = await retrieveUserByRocketChatUserIdAsync(
        read,
        senderId
    );
    if (!senderUser) {
        return;
    }

    await deleteTextMessageInChatThreadAsync(
        http,
        senderUser.teamsUserId,
        messageIdMapping.teamsMessageId,
        messageIdMapping.teamsThreadId,
        senderUserAccessToken
    );
};

export const handlePreFileUploadAsync = async (
    context: IFileUploadContext,
    read: IRead,
    http: IHttp,
    persis: IPersistence,
    modify: IModify
): Promise<void> => {
    const senderRocketChatUserId = context.file.userId;
    const roomId = context.file.rid;
    const fileName = context.file.name;
    const fileMIMEType = context.file.type;
    const fileSize = context.file.size;

    if (fileName.startsWith("thumb-")) {
        // TODO: find a better way to not upload the thumb file for image
        return;
    }

    const isSenderDummerUser = await checkDummyUserByRocketChatUserIdAsync(
        read,
        senderRocketChatUserId
    );
    if (isSenderDummerUser) {
        return;
    }

    const members = await read.getRoomReader().getMembers(roomId);

    const dummyUsers = await findAllDummyUsersInRocketChatUserListAsync(
        read,
        members
    );
    if (dummyUsers && dummyUsers.length > 0) {
        // If there's any dummy user in the room, this is a Teams interop chat room
        // Sanity check has been done in PreMessageSentPrevent for Teams interop scenarios

        // There should be a room record in persist with a bridge user assigned
        const roomRecord = await retrieveRoomByRocketChatRoomIdAsync(
            read,
            roomId
        );
        if (!roomRecord) {
            throw new Error("No room record find for Teams interop room!");
        }

        if (!roomRecord.bridgeUserRocketChatUserId) {
            throw new Error("No bridge user assigned to Teams interop room!");
        }

        const bridgeUser = await retrieveUserByRocketChatUserIdAsync(
            read,
            roomRecord.bridgeUserRocketChatUserId
        );
        let userAccessToken = await retrieveUserAccessTokenAsync(
            read,
            persis,
            roomRecord.bridgeUserRocketChatUserId
        );
        if (!userAccessToken || !bridgeUser) {
            await persistRoomAsync(
                persis,
                roomRecord.rocketChatRoomId,
                roomRecord.teamsThreadId,
                undefined
            );
            throw new Error("Invalid bridge user!");
        }

        const senderUserAccessToken = await retrieveUserAccessTokenAsync(
            read,
            persis,
            senderRocketChatUserId
        );
        if (senderUserAccessToken) {
            // If file uploader already logged in, make the file uploaded by themselves instead of via the bridge user
            userAccessToken = senderUserAccessToken;
        }

        // Upload the file to One Drive
        const uploadFileResponse = await uploadFileToOneDriveAsync(
            http,
            fileName,
            fileMIMEType,
            fileSize,
            context.content,
            userAccessToken
        );

        // Persist file upload record
        if (uploadFileResponse) {
            await persistOneDriveFileAsync(
                persis,
                uploadFileResponse.fileName,
                uploadFileResponse.driveItemId
            );
        }
    }
};

export const handleAddTeamsUserContextualBarSubmitAsync = async (
    operator: IUser,
    room: IRoom,
    teamsUserIdsToSave: string[],
    read: IRead,
    modify: IModify,
    persis: IPersistence,
    http: IHttp,
    app: TeamsBridgeApp
): Promise<void> => {
    const dummyUsersToAdd: UserModel[] = [];
    for (const teamsUserId of teamsUserIdsToSave) {
        const dummyUser = await retrieveDummyUserByTeamsUserIdAsync(
            read,
            teamsUserId
        );
        if (dummyUser) {
            dummyUsersToAdd.push(dummyUser);
        }
    }

    const roomRecord = await retrieveRoomByRocketChatRoomIdAsync(read, room.id);
    if (roomRecord && roomRecord.teamsThreadId) {
        if (!roomRecord.bridgeUserRocketChatUserId) {
            await notifyNotLoggedInUserAsync(
                read,
                operator,
                room,
                app,
                AddUserLoginRequiredHintMessageText
            );
            return;
        }

        // If there's a thread created in Teams side, need to update the participant there as well
        const accessToken = await retrieveUserAccessTokenAsync(
            read,
            persis,
            roomRecord.bridgeUserRocketChatUserId
        );

        const wasSent = await retrieveLoginMessageSentStatus({
            read,
            rocketChatUserId: operator.id,
        });
        if (!accessToken) {
            if (wasSent) {
                return;
            }
            await notifyNotLoggedInUserAsync(
                read,
                operator,
                room,
                app,
                AddUserLoginRequiredHintMessageText
            );
            await saveLoginMessageSentStatus({
                persistence: persis,
                rocketChatUserId: operator.id,
                wasSent: true,
            });
            return;
        }

        for (const dummyUser of dummyUsersToAdd) {
            await addMemberToChatThreadAsync(
                http,
                roomRecord.teamsThreadId,
                dummyUser.teamsUserId,
                accessToken
            );
        }
    }

    const updater = modify.getUpdater();
    const roomBuilder = await updater.room(room.id, operator);

    for (const dummyUser of dummyUsersToAdd) {
        const userToAdd = await read
            .getUserReader()
            .getById(dummyUser.rocketChatUserId);
        if (!userToAdd) {
            console.error("Dummy user to add not found!");
            continue;
        }

        roomBuilder.addMemberToBeAddedByUsername(userToAdd.username);
    }

    await updater.finish(roomBuilder);
};

export const handlePreRoomUserLeaveAsync = async (
    context: IRoomUserLeaveContext,
    read: IRead,
    http: IHttp,
    persistence: IPersistence,
    app: TeamsBridgeApp
): Promise<void> => {
    const roomId = context.room.id;

    const roomRecord = await retrieveRoomByRocketChatRoomIdAsync(read, roomId);
    if (!roomRecord || !roomRecord.teamsThreadId) {
        return;
    }

    const leavingRocketChatUserId = context.leavingUser.id;
    const embeddedLoginUser = await retrieveUserByRocketChatUserIdAsync(
        read,
        leavingRocketChatUserId
    );
    const dummyUser = await retrieveDummyUserByRocketChatUserIdAsync(
        read,
        leavingRocketChatUserId
    );

    if (!embeddedLoginUser && !dummyUser) {
        console.error("Not logged in user or dummy user.");
        return;
    }

    if (!roomRecord.bridgeUserRocketChatUserId) {
        console.error("No bridge user.");
        throw new UserNotAllowedException();
    }

    // If there's a thread created in Teams side, need to update the participant there as well
    const accessToken = await retrieveUserAccessTokenAsync(
        read,
        persistence,
        roomRecord.bridgeUserRocketChatUserId
    );
    if (!accessToken) {
        console.error("No bridge user.");
        await persistRoomAsync(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
        throw new UserNotAllowedException();
    }

    const teamsUserId =
        embeddedLoginUser?.teamsUserId ?? dummyUser?.teamsUserId;
    if (!teamsUserId) {
        return;
    }

    const threadMemberTeamsUserIds = await listMembersInChatThreadAsync(
        http,
        roomRecord.teamsThreadId,
        accessToken
    );
    if (threadMemberTeamsUserIds.find((id) => id === teamsUserId)) {
        await removeMemberFromChatThreadAsync(
            http,
            roomRecord.teamsThreadId,
            teamsUserId,
            accessToken
        );
    }

    if (
        embeddedLoginUser &&
        embeddedLoginUser.teamsUserId === roomRecord.bridgeUserRocketChatUserId
    ) {
        // Clear bridge user if it's been removed
        await persistRoomAsync(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
    }
};

export const handleUserRegistrationAutoRenewAsync = async (
    subscriberEndpointUrl: string,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persis: IPersistence
): Promise<void> => {
    const aadTenantId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadTenantId)
    ).value;
    const aadClientId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadClientId)
    ).value;
    const aadClientSecret = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadClientSecret)
    ).value;

    const allRegistrations = await retrieveAllUserRegistrationsAsync(read);

    if (allRegistrations) {
        for (const registration of allRegistrations) {
            try {
                const refreshToken = await retrieveUserRefreshTokenAsync(
                    read,
                    persis,
                    registration.rocketChatUserId
                );

                if (!refreshToken) {
                    throw new Error(
                        `Refresh token for user ${registration.rocketChatUserId} not found!`
                    );
                }

                const response = await renewUserAccessTokenAsync(
                    http,
                    refreshToken,
                    aadTenantId,
                    aadClientId,
                    aadClientSecret
                );

                const userAccessToken = response.accessToken;

                await persistUserAccessTokenAsync(
                    persis,
                    registration.rocketChatUserId,
                    userAccessToken,
                    response.refreshToken as string,
                    response.expiresIn,
                    response.extExpiresIn
                );

                const subscriptionIds = await listSubscriptionsAsync(
                    http,
                    userAccessToken
                );
                if (subscriptionIds) {
                    for (const subscriptionId of subscriptionIds) {
                        await renewSubscriptionAsync(
                            http,
                            subscriptionId,
                            userAccessToken
                        );
                    }
                } else {
                    const user = await retrieveUserByRocketChatUserIdAsync(
                        read,
                        registration.rocketChatUserId
                    );
                    if (!user) {
                        throw new Error(
                            `User record for user ${registration.rocketChatUserId} not found!`
                        );
                    }

                    await subscribeToAllMessagesForOneUserAsync(
                        http,
                        user.rocketChatUserId,
                        user.teamsUserId,
                        subscriberEndpointUrl,
                        userAccessToken
                    );
                }
            } catch (error) {
                console.error(
                    `Error during renew registration for user ${registration.rocketChatUserId}. Ignore this error and continue. Error: ${error}`
                );
            }
        }
    }
};

const isTeamsMessageAsync = async (
    messageId: string | undefined,
    read: IRead
): Promise<boolean> => {
    if (!messageId) {
        return false;
    }

    const messageIdMapping =
        await retrieveMessageIdMappingByRocketChatMessageIdAsync(
            read,
            messageId
        );
    if (messageIdMapping) {
        return true;
    }

    return false;
};

const findOneTeamsLoggedInUsersAsync = async (
    read: IRead,
    persistence: IPersistence,
    users: IUser[]
): Promise<UserModel | null> => {
    for (const user of users) {
        const accessToken = await retrieveUserAccessTokenAsync(
            read,
            persistence,
            user.id
        );
        if (accessToken) {
            const userModel = await retrieveUserByRocketChatUserIdAsync(
                read,
                user.id
            );
            return userModel;
        }
    }

    return null;
};

const notifyNotLoggedInUserAsync = async (
    read: IRead,
    user: IUser,
    room: IRoom,
    app: TeamsBridgeApp,
    hintMessageText: string
): Promise<void> => {
    const appUser = (await read.getUserReader().getByUsername('microsoftteamsbridge.bot')) as IUser;

    const aadTenantId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadTenantId)
    ).value;
    const aadClientId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadClientId)
    ).value;
    const accessors = app.getAccessors();
    const authEndpointUrl = await getRocketChatAppEndpointUrl(
        accessors,
        AuthenticationEndpointPath
    );
    const loginUrl = getLoginUrl(
        aadTenantId,
        aadClientId,
        authEndpointUrl,
        user.id
    );
    const message = generateHintMessageWithTeamsLoginButton(
        loginUrl,
        appUser,
        room,
        hintMessageText
    );

    await notifyRocketChatUserAsync(message, user, read.getNotifier());
};
