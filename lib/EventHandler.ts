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
    IRoomUserJoinedContext,
    IRoomUserLeaveContext,
    RoomType,
} from "@rocket.chat/apps-engine/definition/rooms";
import { IFileUploadContext } from "@rocket.chat/apps-engine/definition/uploads";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import {
    AddUserLoginRequiredHintMessageText,
    AuthenticationEndpointPath,
    BridgeUserNotificationMessageText,
    DefaultThreadName,
    LoggedInBridgeUserRequiredHintMessageText,
    LoginRequiredHintMessageText,
    UnsupportedScenarioHintMessageText,
} from "./Const";
import {
    combineRocketChatMessagesToTeamsMessage,
    generateHintMessageWithTeamsLoginButton,
    isBridgedMessageFormat,
    mapRocketChatMessageToTeamsMessageV2,
    notifyRocketChatUserAsync,
    notifyRocketChatUserInRoomAsync,
} from "./MessageHelper";
import {
    addMemberToChatThreadAsync,
    createChatThreadAsync,
    createOneOnOneChatThreadAsync,
    deleteAllSubscriptions,
    deleteTextMessageInChatThreadAsync,
    listMembersInChatThreadAsync,
    removeMemberFromChatThreadAsync,
    sendFileMessageToChatThreadAsync,
    sendTextMessageToChatThreadAsync,
    shareOneDriveFileAsync,
    subscribeToAllMessagesForOneUserAsync,
    updateTextMessageInChatThreadAsync,
    uploadFileToOneDriveAsync,
} from "./MicrosoftGraphApi";
import { LoginMessage, MessageMapping, OneDriveFile, Room, UploadMapping, UserMapping, UserRegistration } from "./PersistHelper";
import type { UploadMappingModel, UserModel } from "./PersistHelper";
import { getLoginUrl, getNotificationEndpointUrl, getRocketChatAppEndpointUrl } from "./UrlHelper";
import { getAllUsersAccessTokensAsync, getUserAccessTokenAsync } from "./AuthHelper";
import { PreventRegistry } from "./PreventRegistry";

export const handlePreMessageSentPreventAsync = async (options: {
    message: IMessage;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp,
}): Promise<boolean> => {
    const { message, read, persistence, app, http } = options;
    try {
        const appUser = await read.getUserReader().getAppUser(app.getID()) as IUser;
        const notifier = read.getNotifier();

        if (message.threadId) {
            const isTeamsMessageThread = await isTeamsMessageAsync(
                message.threadId,
                read
            );
            if (isTeamsMessageThread) {
                // There's no thread message concept in Teams
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
            const messageMapping = await MessageMapping.findByRCMessageId(read, message.id as string);
            if (messageMapping?.teamsMessageId) {
                return true;
            }

            if (!await Room.isBridged(read, message.room.id)) {
                return false;
            }

            const members = await read.getRoomReader().getMembers(message.room.id);
let roomRecord = await Room.findByRCRoomId(
                read,
                message.room.id
            );

            if (roomRecord) {
                if (roomRecord.bridgeUserRocketChatUserId) {
                    const accessToken = await getUserAccessTokenAsync({
                        read,
                        persistence,
                        rocketChatUserId: roomRecord.bridgeUserRocketChatUserId,
                        app,
                        http,
                    });
                    if (!accessToken) {
                        roomRecord.bridgeUserRocketChatUserId = undefined;
                    }
                }
            } else {
                roomRecord = { rocketChatRoomId: message.room.id };
            }

            if (!roomRecord.bridgeUserRocketChatUserId) {
                const loggedInUser = await findOneTeamsLoggedInUsersAsync({
                    read,
                    persistence,
                    users: members,
                    app,
                    http,
                });
                const isOneOnOneDirectMessage =
                    roomType === RoomType.DIRECT_MESSAGE && members.length === 2;
                if (loggedInUser) {
                    roomRecord.bridgeUserRocketChatUserId = loggedInUser.rocketChatUserId;
                    if (!isOneOnOneDirectMessage) {
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
                    }
                } else {
                    const wasSent = await LoginMessage.get({
                        read,
                        rocketChatUserId: message.sender.id,
                    });
                    if (!wasSent) {
                        await notifyNotLoggedInUserAsync(
                            read,
                            message.sender,
                            message.room,
                            app,
                            isOneOnOneDirectMessage
                                ? LoginRequiredHintMessageText
                                : LoggedInBridgeUserRequiredHintMessageText
                        );
                        await LoginMessage.save({
                            persistence,
                            rocketChatUserId: message.sender.id,
                            wasSent: true,
                        });
                    }
                }
            }

            await Room.persist(
                persistence,
                roomRecord.rocketChatRoomId,
                roomRecord.teamsThreadId,
                roomRecord.bridgeUserRocketChatUserId
            );
        }
        return false;
    } catch (error) {
        app.getLogger().error(error);
        return false;
    }
};

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
            textMessage: messageText,
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

export const handlePreMessageOperationPreventAsync = async (_options: {
    message: IMessage,
    read: IRead,
    persistence: IPersistence,
    app: TeamsBridgeApp,
    http: IHttp,
}): Promise<boolean> => {
    // Single-bot architecture: no per-user Teams identity to check.
    // Edit/delete operations on bridged messages are controlled by the
    // message-ID mapping check in the update/delete handlers themselves.
    return false;
};

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
        const { text, attachments } =  await mapRocketChatMessageToTeamsMessageV2({
            message,
            read,
            http,
            accessToken: senderUserAccessToken,
            messageIdMapping,
        })
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

export const handlePostMessageDeletedAsync = async (options: {
    message: IMessage;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp;
}): Promise<void> => {
    const { message, read, persistence, app, http } = options;
    if (
        await PreventRegistry.capture(
            persistence,
            `PreventPostMessageDeleteHook/${message.id}`
        )
    ) {
        // Prevent duplicate processing
        return;
    }

    const msgId = message.id;
    if (!msgId) {
        return;
    }

    // --- Step 1: Resolve mappings (message ↔ upload ↔ teams) ---
    const {
        messageIdMapping,
        uploadMappings,
        currentUploadMapping,
        mainMessage,
    } = await resolveMappings(read, { ...message, id: msgId });

    if (
        !messageIdMapping &&
        !currentUploadMapping &&
        uploadMappings.length === 0
    ) {
        return;
    }

    // --- Step 2: Ensure sender info (user + access token) ---
    const { senderUser, accessToken } = await ensureSenderInfo({
        senderId: message.sender.id,
        read,
        persistence,
        app,
        http,
    });

    if (!senderUser || !accessToken) {
        return;
    }

    // --- Step 3: Clean up mappings in persistence ---
    if (currentUploadMapping) {
        await UploadMapping.delete({
            persistence,
            rocketchatUploadId: currentUploadMapping.rocketchatUploadId,
            teamsMessageId: currentUploadMapping.teamsMessageId,
        });
    }

    if (messageIdMapping?.rocketChatMessageId === msgId) {
        await MessageMapping.delete({ persistence, ...messageIdMapping });
    }

    // --- Step 4: Prepare Teams update ---
    const teamsIds = {
        messageId:
            messageIdMapping?.teamsMessageId ||
            currentUploadMapping?.teamsMessageId,
        threadId:
            messageIdMapping?.teamsThreadId ||
            currentUploadMapping?.teamsThreadId,
    };

    if (!teamsIds.messageId || !teamsIds.threadId) {
        return;
    }

    const deletedIds = {
        messages: new Set([msgId]),
        uploads: message.file?._id ? new Set<string>([message.file._id]) : new Set<string>(),
    };

    const isBridge = isBridgedMessageFormat(mainMessage?.text || "");
    const { text, shouldDeleteTeamsMessage, attachments } =
        await combineRocketChatMessagesToTeamsMessage({
            read,
            messages: mainMessage ? [mainMessage] : [],
            messageIdMapping: {
                rocketChatMessageId: msgId,
                teamsMessageId: teamsIds.messageId,
                teamsThreadId: teamsIds.threadId,
            },
            deletedMessages: deletedIds.messages,
            deletedUploads: deletedIds.uploads,
            forceBridgedMessage: isBridge,
            originalSenderName: isBridge
                ? message.sender.name || message.sender.username
                : undefined,
            uploadMappings,
            http,
            accessToken,
        });

    // --- Step 5: Execute Teams update/delete ---
    if (shouldDeleteTeamsMessage) {
        await PreventRegistry.set(
            persistence,
            `PreventPostMessageDeleteHook/${message.id}`
        );
        await deleteTextMessageInChatThreadAsync(
            http,
            senderUser.teamsUserId,
            teamsIds.messageId,
            teamsIds.threadId,
            accessToken
        );
    } else {
        await PreventRegistry.set(
            persistence,
            `PreventPostMessageUpdateHook/${message.id}`
        );
        await updateTextMessageInChatThreadAsync({
            http,
            textMessage: text,
            messageType: "html",
            messageId: teamsIds.messageId,
            threadId: teamsIds.threadId,
            userAccessToken: accessToken,
            attachments,
        });
    }
};

/* ---------------------- Helpers ----------------------- */

async function resolveMappings(read: IRead, message: IMessage & { id: string }) {
    let mainMessage: IMessage | null = null;
    let messageIdMapping =
        await MessageMapping.findByRCMessageId(
            read,
            message.id
        );

    let uploadMappings: UploadMappingModel[] = [];
    if (message.file?._id) {
        uploadMappings =
            await UploadMapping.findAllByRCUploadId(
                read,
                message.file._id
            );
    }

    const currentUploadMapping = uploadMappings.find(
        (u) => u.rocketchatUploadId === message.file?._id
    );

    if (currentUploadMapping && !messageIdMapping) {
        messageIdMapping = await MessageMapping.findByTeamsMessageId(
            read,
            currentUploadMapping.teamsMessageId
        );
        if (messageIdMapping) {
            mainMessage =
                (await read
                    .getMessageReader()
                    .getById(messageIdMapping.rocketChatMessageId)) || null;
        }
    } else if (!currentUploadMapping && messageIdMapping) {
        uploadMappings = await UploadMapping.findByTeamsMessageId(
            read,
            messageIdMapping.teamsMessageId
        );
        mainMessage = message;
    }

    return {
        messageIdMapping,
        uploadMappings,
        currentUploadMapping,
        mainMessage,
    };
}

async function ensureSenderInfo({
    senderId,
    read,
    persistence,
    app,
    http,
}: {
    senderId: string;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp;
}) {
    const [accessToken, senderUser] = await Promise.all([
        getUserAccessTokenAsync({
            read,
            persistence,
            rocketChatUserId: senderId,
            app,
            http,
        }),
        UserMapping.findByRCUserId(read, senderId),
    ]);

    return { senderUser, accessToken };
}


export const handlePreFileUploadAsync = async (options: {
    context: IFileUploadContext;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp;
}): Promise<void> => {
    const { context, app, http, persistence, read} = options;
    const senderRocketChatUserId = context.file.userId;
    const roomId = context.file.rid;
    const fileName = context.file.name;
    const fileMIMEType = context.file.type;
    const fileSize = context.file.size;

    if (fileName.startsWith("thumb-")) {
        // TODO: find a better way to not upload the thumb file for image
        return;
    }

    // Skip uploads made by the app bot itself (e.g. inbound relayed files)
    const appUser = await read.getUserReader().getAppUser(app.getID());
    if (appUser && senderRocketChatUserId === appUser.id) {
        return;
    }

    if (!await Room.isBridged(read, roomId)) {
        return;
    }

    // There should be a room record in persist with a bridge user assigned
    const roomRecord = await Room.findByRCRoomId(read, roomId);
    if (!roomRecord) {
        throw new Error("No room record find for Teams interop room!");
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

    const senderUserAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: senderRocketChatUserId,
        app,
        http,
    });
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
        await OneDriveFile.persist(
            persistence,
            uploadFileResponse.fileName,
            uploadFileResponse.driveItemId,
        );
    }
};

export const handleAddTeamsUserContextualBarSubmitAsync = async (options: {
    operator: IUser;
    room: IRoom;
    teamsUserIdsToSave: string[];
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const {
        app,
        http,
        operator,
        persistence,
        read,
        room,
        teamsUserIdsToSave,
    } = options;

    const roomRecord = await Room.findByRCRoomId(read, room.id);
    if (!roomRecord?.teamsThreadId) {
        return;
    }

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

    // If there's a thread created in Teams side, update the participants there as well
    const accessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: roomRecord.bridgeUserRocketChatUserId,
        app,
        http,
    });

    if (!accessToken) {
        const wasSent = await LoginMessage.get({
            read,
            rocketChatUserId: operator.id,
        });
        if (!wasSent) {
            await notifyNotLoggedInUserAsync(
                read,
                operator,
                room,
                app,
                AddUserLoginRequiredHintMessageText
            );
            await LoginMessage.save({
                persistence,
                rocketChatUserId: operator.id,
                wasSent: true,
            });
        }
        return;
    }

    // Single-bot architecture: iterate directly over Teams user IDs —
    // no dummy RC users to look up or add to the RC room.
    for (const teamsUserId of teamsUserIdsToSave) {
        await addMemberToChatThreadAsync(
            http,
            roomRecord.teamsThreadId,
            teamsUserId,
            accessToken
        );
    }
};

export const handlePostRoomUserJoinedAsync = async (options: {
    context: IRoomUserJoinedContext;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const { context, read, persistence, app } = options;
    const { joiningUser, room } = context;

    const appUser = await read.getUserReader().getAppUser(app.getID());
    if (!appUser || joiningUser.id !== appUser.id) {
        return;
    }

    // setBridgeRoomActiveAsync preserves any existing teamsThreadId /
    // bridgeUserRocketChatUserId, so re-adding the bot reuses the same thread
    await Room.setBridgeActive(persistence, read, room.id, true);

    app.getLogger().info(
        `[TeamsBridge] Room "${room.displayName || room.id}" is now an active bridge room ` +
        `(app user added by ${context.inviter?.username ?? 'unknown'}).`
    );
};

export const handlePreRoomUserLeaveAsync = async (options: {
    context: IRoomUserLeaveContext;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const { app, context, http, persistence, read} = options;
    const roomId = context.room.id;
    const leavingRocketChatUserId = context.leavingUser.id;

    // When the app bot is removed, pause bridging without touching the Teams thread.
    // Re-adding the bot later will reactivate the same thread.
    const appUser = await read.getUserReader().getAppUser(app.getID());
    if (appUser && leavingRocketChatUserId === appUser.id) {
        await Room.setBridgeActive(persistence, read, roomId, false);
        app.getLogger().info(`[TeamsBridge] Room "${context.room.displayName || roomId}" bridging paused (app user removed).`);
        return;
    }

    const roomRecord = await Room.findByRCRoomId(read, roomId);
    if (!roomRecord || !roomRecord.teamsThreadId) {
        return;
    }

    const embeddedLoginUser = await UserMapping.findByRCUserId(
        read,
        leavingRocketChatUserId
    );

    if (!embeddedLoginUser) {
        return;
    }

    if (!roomRecord.bridgeUserRocketChatUserId) {
        console.error("No bridge user.");
        throw new UserNotAllowedException();
    }

    const accessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: roomRecord.bridgeUserRocketChatUserId,
        app,
        http,
    });
    if (!accessToken) {
        console.error("No bridge user.");
        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
        throw new UserNotAllowedException();
    }

    const teamsUserId = embeddedLoginUser.teamsUserId;
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
        embeddedLoginUser.teamsUserId === roomRecord.bridgeUserRocketChatUserId
    ) {
        // Clear bridge user if it's been removed
        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
    }
};

export const handleUserRegistrationAutoRenewAsync = async (options: {
    subscriberEndpointUrl: string;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp,
}): Promise<void> => {
    const { http, persistence, read, subscriberEndpointUrl, app } = options;

    const allRegistrations = await UserRegistration.findAll(read);

    if (allRegistrations) {
        const errorUserIds: string[] = [];
        for (const registration of allRegistrations) {
            try {
                const userAccessToken = await getUserAccessTokenAsync({
                    app,
                    http,
                    persistence,
                    read,
                    rocketChatUserId: registration.rocketChatUserId,
                });

                if (!userAccessToken) {
                    errorUserIds.push(registration.rocketChatUserId);
                    continue;
                }

                const user = await UserMapping.findByRCUserId(
                    read,
                    registration.rocketChatUserId
                );

                if (!user) {
                    throw new Error(
                        `User record for user ${registration.rocketChatUserId} not found!`
                    );
                }

                await subscribeToAllMessagesForOneUserAsync({
                    read,
                    http,
                    persis: persistence,
                    rocketChatUserId: user.rocketChatUserId,
                    subscriberEndpointUrl,
                    teamsUserId: user.teamsUserId,
                    userAccessToken,
                    renewIfExists: true,
                    forceRenew: false,
                });
            } catch (error) {
                console.error(
                    `Error during renew registration for user ${registration.rocketChatUserId}. Ignore this error and continue. Error: ${error}`
                );
            }
        }
        if (errorUserIds.length) {
            app.getLogger().error(`Could not refresh user access token for users: ${errorUserIds.join(', ')}`)
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
        await MessageMapping.findByRCMessageId(
            read,
            messageId
        );
    if (messageIdMapping) {
        return true;
    }

    return false;
};

const findOneTeamsLoggedInUsersAsync = async (options: {
    read: IRead;
    persistence: IPersistence;
    users: IUser[];
    app: TeamsBridgeApp,
    http: IHttp,
}): Promise<UserModel | null> => {
    const { app, http, persistence, read, users } = options;
    for (const user of users) {
        const accessToken = await getUserAccessTokenAsync({
            read,
            persistence,
            rocketChatUserId: user.id,
            app,
            http,
        });
        if (accessToken) {
            const userModel = await UserMapping.findByRCUserId(
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

const deleteAllUsersSubscriptions = async (options: {
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
}) => {
    const { read, persistence, http, app } = options;
    const allRegisteredUsers = await getAllUsersAccessTokensAsync({
        read,
        http,
        app,
        persistence,
    });

    if (!allRegisteredUsers) {
        return;
    }

    const batchSize = 10;
    for (let i = 0; i < allRegisteredUsers.length; i += batchSize) {
        const batch = allRegisteredUsers.slice(i, i + batchSize);
        const deletePromises = batch.map(async ({ accessToken, rocketChatUserId }) => {
            try {
                const notificationUrl = await getNotificationEndpointUrl({
                    appAccessors: app.getAccessors(),
                    rocketChatUserId: rocketChatUserId,
                });
                if (accessToken) {
                    await deleteAllSubscriptions(
                        http,
                        accessToken,
                        notificationUrl
                    );
                }
            } catch (error) {
                console.error(
                    `Error deleting subscriptions for user: ${error.message}`
                );
            }
        });
        await Promise.all(deletePromises); // Wait for the current batch to complete
    }
};

export const handleUninstallApp = async (options: {
    read: IRead;
    http: IHttp;
    modify: IModify;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}) => {
    const { modify, app } = options;
    try {
        await deleteAllUsersSubscriptions(options),
        await app.deleteAppUsers(modify);
    } catch (error) {
        console.error(`Error during app uninstallation: ${error.message}`);
    }
};
