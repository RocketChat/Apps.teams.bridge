import {
    IHttp,
    IModify,
    INotifier,
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
    BridgeUserNotificationMessageText,
    LoggedInBridgeUserRequiredHintMessageText,
    LoginRequiredHintMessageText,
    UnsupportedScenarioHintMessageText
} from "./Const";
import {
    generateHintMessageWithTeamsLoginButton,
    mapRocketChatMessageToTeamsMessage,
    notifyRocketChatUserAsync,
    notifyRocketChatUserInRoomAsync
} from "./MessageHelper";
import {
    createChatThreadAsync,
    createOneOnOneChatThreadAsync,
    deleteTextMessageInChatThreadAsync,
    sendTextMessageToChatThreadAsync,
    updateTextMessageInChatThreadAsync
} from "./MicrosoftGraphApi";
import {
    checkDummyUserByRocketChatUserIdAsync,
    persistMessageIdMappingAsync,
    persistRoomAsync,
    retrieveDummyUserByRocketChatUserIdAsync,
    retrieveMessageIdMappingByRocketChatMessageIdAsync,
    retrieveRoomByRocketChatRoomIdAsync,
    retrieveUserAccessTokenAsync,
    retrieveUserByRocketChatUserIdAsync,
    UserModel
} from "./PersistHelper";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "./UrlHelper";

export const handlePreMessageSentPreventAsync = async (
    message: IMessage,
    read: IRead,
    persistence: IPersistence,
    app: TeamsBridgeApp) : Promise<boolean> => {
    const appUser = (await read.getUserReader().getAppUser()) as IUser;
    const notifier = read.getNotifier();

    if (message.threadId) {
        const isTeamsMessageThread = await isTeamsMessageAsync(message.threadId, read);
        if (isTeamsMessageThread) {
            // There's no thread message concept in Teams
            // Thread message is not a supported scenario for Teams interop

            await notifyRocketChatUserInRoomAsync(
                UnsupportedScenarioHintMessageText("Thread Message"),
                appUser,
                message.sender,
                message.room,
                notifier);

            return true;
        }
    }

    let preventOperation = false;
    const roomType = message.room.type;
    if (roomType === RoomType.PRIVATE_GROUP || roomType === RoomType.DIRECT_MESSAGE) {
        // If room type is PRIVATE_GROUP or DIRECT_MESSAGE, check if there's any dummy user in the room
        const members = await read.getRoomReader().getMembers(message.room.id);

        const dummyUsers = await findAllDummyUsersAsync(read, members);
        console.log(dummyUsers);

        if (dummyUsers && dummyUsers.length > 0) {
            // If there are dummy users in the room, check whether there's at least one teams-logged in user in this room

            // Find whether there's an existing room record
            let roomRecord = await retrieveRoomByRocketChatRoomIdAsync(read, message.room.id);

            console.log("VVV==roomRecord==VVV");
            console.log(roomRecord);
            console.log("^^^==roomRecord==^^^");

            if (roomRecord) {
                // If there's an existing room record, check whether it has a bridge user
                if (roomRecord.bridgeUserRocketChatUserId) {
                    // If this room already has assigned a bridge user, check the bridge user login status

                    const accessToken = await retrieveUserAccessTokenAsync(read, roomRecord.bridgeUserRocketChatUserId);
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
                const loggedInUser = await findOneTeamsLoggedInUsersAsync(read, members);
                const isOneOnOneDirectMessage = roomType === RoomType.DIRECT_MESSAGE && members.length === 2;
                if (loggedInUser) {
                    // Assign the room a bridge user
                    roomRecord.bridgeUserRocketChatUserId = loggedInUser.rocketChatUserId;
    
                    // For 1:1 dm chat, no further action required
                    if (!isOneOnOneDirectMessage) {
                        // For other type of chat room
                        // Notify the bridge user that he has became the bridge of this room
                        // All messages sent by unlogged in user will be delivered to Microsoft Teams by him
                        const bridgeUser = await read.getUserReader().getById(loggedInUser.rocketChatUserId);
                        await notifyRocketChatUserInRoomAsync(BridgeUserNotificationMessageText, appUser, bridgeUser, message.room, notifier);

                        // TODO: if there's an existing Teams thread, update thread topic

                        // TODO: send a message to Microsoft Teams to let the user there know the bridge user represents some other users
                    }
                } else {
                    // If there's no logged in user in the room, prevent the message
                    preventOperation = true;
    
                    if (isOneOnOneDirectMessage) {
                        // For 1:1 chat, notify the sender to login
                        await notifyNotLoggedInUserAsync(read, message.sender, message.room, app, LoginRequiredHintMessageText);
                    } else {
                        // For other type of chat room
                        // Notify the message sender there's no available bridge user
                        await notifyNotLoggedInUserAsync(read, message.sender, message.room, app, LoggedInBridgeUserRequiredHintMessageText);
                    }
                }
            }

            // Persist the room record
            await persistRoomAsync(
                persistence,
                roomRecord.rocketChatRoomId,
                roomRecord.teamsThreadId,
                roomRecord.bridgeUserRocketChatUserId);
        }
    }

    return preventOperation;
};

export const handlePostMessageSentAsync = async (
    message: IMessage,
    read: IRead,
    http: IHttp,
    persistence: IPersistence) : Promise<void> => {

    //console.log("VVV==message==VVV");
    //console.log(message);
    //console.log("^^^==message==^^^");
    const isSenderDummyUser = await checkDummyUserByRocketChatUserIdAsync(read, message.sender.id);
    if (isSenderDummyUser) {
        console.log('Message sender is a dummy user, stop processing.');
        return;
    }

    const roomId = message.room.id;
    const members = await read.getRoomReader().getMembers(roomId);

    const dummyUsers = await findAllDummyUsersAsync(read, members);
    console.log(dummyUsers);
    if (dummyUsers && dummyUsers.length > 0) {
        // If there's any dummy user in the room, this is a Teams interop chat room
        // Sanity check has been done in PreMessageSentPrevent for Teams interop scenarios

        // There should be a room record in persist with a bridge user assigned
        const roomRecord = await retrieveRoomByRocketChatRoomIdAsync(read, roomId);
        if (!roomRecord) {
            throw new Error('No room record find for Teams interop room!');
        }

        if (!roomRecord.bridgeUserRocketChatUserId) {
            throw new Error('No bridge user assigned to Teams interop room!');
        }

        const bridgeUser = await retrieveUserByRocketChatUserIdAsync(read, roomRecord.bridgeUserRocketChatUserId);
        let userAccessToken = await retrieveUserAccessTokenAsync(read, roomRecord.bridgeUserRocketChatUserId);
        if (!userAccessToken || !bridgeUser) {
            await persistRoomAsync(
                persistence,
                roomRecord.rocketChatRoomId,
                roomRecord.teamsThreadId,
                undefined);
            throw new Error('Invalid bridge user!');
        }

        if (!roomRecord.teamsThreadId) {
            // Not yet a thread exist in Teams side, create one & persist in room record
            if (message.room.type === RoomType.DIRECT_MESSAGE && members.length === 2) {
                // If 1:1 DM, create 1:1 Teams chat thread
                const response = await createOneOnOneChatThreadAsync(http, bridgeUser.teamsUserId, dummyUsers[0].teamsUserId, userAccessToken);
                roomRecord.teamsThreadId = response.threadId;
            } else {
                // If other room type, create Teams group chat thread
                const teamsIds: string[] = [];
                for (const member of members) {
                    const user = await retrieveUserByRocketChatUserIdAsync(read, member.id);
                    if (user) {
                        teamsIds.push(user.teamsUserId);
                    }
                }

                for (const dummyUser of dummyUsers) {
                    teamsIds.push(dummyUser.teamsUserId);
                }

                const bridgeUserName = (await read.getUserReader().getById(bridgeUser.rocketChatUserId)).name;
                const response = await createChatThreadAsync(http, teamsIds, bridgeUserName, userAccessToken);
                roomRecord.teamsThreadId = response.threadId;
            }
            
            await persistRoomAsync(
                persistence,
                roomRecord.rocketChatRoomId,
                roomRecord.teamsThreadId,
                roomRecord.bridgeUserRocketChatUserId);
        }

        let messageText = message.text;
        if (!messageText) {
            messageText = '';
        }

        const isMessageBridged = bridgeUser.rocketChatUserId !== message.sender.id;
        let originalSenderName = isMessageBridged ? message.sender.name : undefined;

        const senderUserAccessToken = await retrieveUserAccessTokenAsync(read, message.sender.id);
        if (senderUserAccessToken) {
            // If message sender already logged in, make the message sent by themselves instead of via the bridge user
            userAccessToken = senderUserAccessToken;
            originalSenderName = undefined;
        }

        // Mapping message content format
        messageText = mapRocketChatMessageToTeamsMessage(messageText, originalSenderName);

        // Send the message to the chat thread
        const response = await sendTextMessageToChatThreadAsync(
            http,
            messageText,
            roomRecord.teamsThreadId,
            userAccessToken);

        const teamsMessageId = response.messageId;
        const rocketChatMessageId = message.id as string;

        await persistMessageIdMappingAsync(persistence, rocketChatMessageId, teamsMessageId, roomRecord.teamsThreadId);
    }
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

const findAllDummyUsersAsync = async (read: IRead, users: IUser[]) : Promise<UserModel[]> => {
    console.log("findAllDummyUsersAsync");

    const result : UserModel[] = [];
    for (const user of users) {
        const userModel = await retrieveDummyUserByRocketChatUserIdAsync(read, user.id);
        if (userModel) {
            result.push(userModel);
        }
    }
    
    console.log(`Find ${result.length} dummy users`);

    return result;
};

const findOneTeamsLoggedInUsersAsync = async (read: IRead, users: IUser[]) : Promise<UserModel | null> => {
    console.log("findAllTeamsLoggedInUsersAsync");

    for (const user of users) {
        const accessToken = await retrieveUserAccessTokenAsync(read, user.id);
        if (accessToken) {
            const userModel = await retrieveUserByRocketChatUserIdAsync(read, user.id);
            return userModel;
        }
    }
    
    console.log('Not find any Teams logged in users!');

    return null;
};

const notifyNotLoggedInUserAsync = async (
    read: IRead,
    user: IUser,
    room: IRoom,
    app: TeamsBridgeApp,
    hintMessageText: string) : Promise<void> => {
    const appUser = (await read.getUserReader().getAppUser()) as IUser;

    const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
    const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
    const accessors = app.getAccessors();
    const authEndpointUrl = await getRocketChatAppEndpointUrl(accessors, AuthenticationEndpointPath);
    const loginUrl = getLoginUrl(aadTenantId, aadClientId, authEndpointUrl, user.id);
    const message = generateHintMessageWithTeamsLoginButton(loginUrl, appUser, room, hintMessageText)
    
    await notifyRocketChatUserAsync(message, user, read.getNotifier());
};
