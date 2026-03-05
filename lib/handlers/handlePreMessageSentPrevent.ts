import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { UnsupportedScenarioHintMessageText } from "../Const";
import { notifyRocketChatUserInRoomAsync } from "../Notifier";
import { MessageMapping, Room } from "../PersistHelper";
import { PreventRegistry } from "../PreventRegistry";

export const handlePreMessageSentPreventAsync = async (options: {
    message: IMessage;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp,
}): Promise<boolean> => {
    const { message, read, app } = options;
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
