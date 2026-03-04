import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import {
    BridgeUserNotificationMessageText,
    LoggedInBridgeUserRequiredHintMessageText,
    LoginRequiredHintMessageText,
    UnsupportedScenarioHintMessageText,
} from "../Const";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { notifyNotLoggedInUserAsync, notifyRocketChatUserInRoomAsync } from "../Notifier";
import { LoginMessage, MessageMapping, Room, UserMapping } from "../PersistHelper";
import type { UserModel } from "../PersistHelper";
import { PreventRegistry } from "../PreventRegistry";

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

const isTeamsMessageAsync = async (
    messageId: string | undefined,
    read: IRead
): Promise<boolean> => {
    if (!messageId) {
        return false;
    }

    const messageIdMapping = await MessageMapping.findByRCMessageId(read, messageId);
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
            const userModel = await UserMapping.findByRCUserId(read, user.id);
            return userModel;
        }
    }

    return null;
};
