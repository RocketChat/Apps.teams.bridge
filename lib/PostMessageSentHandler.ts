import { IHttp, IModify, IPersistence, IRead } from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom, RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { nofityRocketChatUserInRoomAsync } from "./Messages";
import { checkDummyUserAsync, retrieveDummyUserAsync, retrieveUserAccessTokenAsync } from "./PersistHelper";

export const handlePostMessageSentAsync = async (
    message: IMessage,
    read: IRead,
    http: IHttp,
    persistence: IPersistence,
    modify: IModify) : Promise<void> => {
    // In first version, we'll only support sending text message in 1:1 chat with Teams user

    // If the message is not a text message, stop processing
    if (!isTextMessage(message)) {
        return;
    }

    // If the message is not sent in a 1:1 chat, stop processing
    const room = message.room;
    if (!isOneOnOneChat(room)) {
        return;
    }

    // If the message receiver is not a Teams Dummy user, stop processing
    const receiverId = getDirectMessageReceiverId(message, room);
    if (receiverId === undefined) {
        return;
    }
    const isReceiverDummyUser = await checkDummyUserAsync(read, receiverId);
    if (!isReceiverDummyUser) {
        return;
    }

    // If the message sender has not logged in to Teams
    // Send a notification to let the sender know he need to logged in to Teams to start cross platform collaboration
    const senderId = message.sender.id;
    const senderUserAccessToken = await retrieveUserAccessTokenAsync(read, senderId);
    if (senderUserAccessToken === null) {
        await notifyNotLoggedInUserAsync(read, message.sender, room, modify);
        return;
    }

    // All checks passed, find the dummy user record
    const dummyUser = await retrieveDummyUserAsync(read, receiverId);
};

const notifyNotLoggedInUserAsync = async (read: IRead, user: IUser, room: IRoom, modify: IModify) : Promise<void> => {
    const errorHintMessage = 'The Rocket.Chat user you are messaging represents a colleague in your organization using Microsoft Teams. '
    + 'The message can NOT be delivered to the user on Microsoft Teams before you start cross platform collaboration for your account. '
    + 'For details, see https://github.com/RocketChat/Apps.teams.bridge/blob/main/docs/support.md.';
    const appUser = (await read.getUserReader().getAppUser()) as IUser;
    await nofityRocketChatUserInRoomAsync(errorHintMessage, appUser, user, room, modify);
};

const getDirectMessageReceiverId = (message: IMessage, room: IRoom) : string | undefined => {
    return room.userIds?.find(id => id !== message.sender.id);
};

const isOneOnOneChat = (room: IRoom) : boolean => {
    if (room.type === RoomType.DIRECT_MESSAGE && room.userIds && room.userIds.length === 2) {
        return true;
    }

    return false;
};

const isTextMessage = (message: IMessage) : boolean => {
    if (message.text) {
        return true;
    }

    return false;
};
