import {
    IMessageBuilder,
    IModify,
    IModifyCreator,
    INotifier,
    IRead,
    IRoomBuilder,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom, RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";

export const sendRocketChatOneOnOneMessageAsync = async (
    message: string,
    sender: IUser,
    receiver: IUser,
    read: IRead,
    modify: IModify) : Promise<void> => {
    const creator: IModifyCreator = modify.getCreator();
    const roomBuilder: IRoomBuilder = creator
        .startRoom()
        .setCreator(sender)
        .setType(RoomType.DIRECT_MESSAGE)
        .setSlugifiedName(`dm_${sender.username}_${receiver.username}`)
        .addMemberToBeAddedByUsername(sender.username)
        .addMemberToBeAddedByUsername(receiver.username);

    const roomId = await creator.finish(roomBuilder);
    const room = (await read.getRoomReader().getById(roomId)) as IRoom;

    const messageTemplate: IMessage = {
        text: message,
        sender: sender,
        room
    };

    const messageBuilder: IMessageBuilder = creator.startMessage(messageTemplate);
    await creator.finish(messageBuilder);
};

export const nofityRocketChatUserInRoomAsync = async (
    message: string,
    appUser: IUser,
    user: IUser,
    room: IRoom,
    modify: IModify) : Promise<void> => {
    const notifier: INotifier = modify.getNotifier();

    const messageTemplate: IMessage = {
        text: message,
        sender: appUser,
        room
    };
    await notifier.notifyUser(user, messageTemplate);
};
