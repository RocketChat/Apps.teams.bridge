import {
    IMessageBuilder,
    IModify,
    IModifyCreator,
    INotifier,
    IRead,
    IRoomBuilder,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage, IMessageAction, IMessageAttachment, MessageActionType } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom, RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { shortnameToUnicode } from "emojione";
import { LoginButtonText } from "./Const";
import { MessageContentType } from "./MicrosoftGraphApi";

export const sendRocketChatOneOnOneMessageAsync = async (
    message: string,
    sender: IUser,
    receiver: IUser,
    read: IRead,
    modify: IModify) : Promise<string> => {
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
    return await creator.finish(messageBuilder);
};

export const sendRocketChatMessageInRoomAsync = async (
    message: string,
    sender: IUser,
    room: IRoom,
    modify: IModify) : Promise<string> => {
    const creator: IModifyCreator = modify.getCreator();

    const messageTemplate: IMessage = {
        text: message,
        sender: sender,
        room
    };

    const messageBuilder: IMessageBuilder = creator.startMessage(messageTemplate);
    return await creator.finish(messageBuilder);
};

export const notifyRocketChatUserInRoomAsync = async (
    message: string,
    appUser: IUser,
    user: IUser,
    room: IRoom,
    notifier: INotifier) : Promise<void> => {
    const messageTemplate: IMessage = {
        text: message,
        sender: appUser,
        room
    };

    await notifyRocketChatUserAsync(messageTemplate, user, notifier);
};

export const notifyRocketChatUserAsync = async (
    message: IMessage,
    user: IUser,
    notifier: INotifier) : Promise<void> => {
    await notifier.notifyUser(user, message);
};

export const generateHintMessageWithTeamsLoginButton = (
    loginUrl: string,
    sender: IUser,
    room: IRoom,
    hintMessageText: string) : IMessage  =>{
    const buttonAction: IMessageAction = {
        type: MessageActionType.BUTTON,
        text: LoginButtonText,
        url: loginUrl,
    };

    const buttonAttachment: IMessageAttachment = {
        actions: [
            buttonAction
        ]
    };

    const message: IMessage = {
        text: hintMessageText,
        sender: sender,
        room,
        attachments: [
            buttonAttachment
        ]
    };

    return message;
};

export const mapTeamsMessageToRocketChatMessage = (
    teamsMessage: string,
    contentType: MessageContentType | undefined) : string => {
    let rocketChatMessage = teamsMessage;
    if (contentType && contentType === MessageContentType.Html) {
        // TODO: find a better way to trim html tag from html messages
        const nbspPattern = /&nbsp;/g;
        const htmpTagPattern = /<\/?[^>]+>/g;

        rocketChatMessage = rocketChatMessage.replace(nbspPattern, ' ');
        rocketChatMessage = rocketChatMessage.replace(htmpTagPattern, match => {
            if (match.indexOf('itemtype="http://schema.skype.com/Emoji"') > 0) {
                const index = match.indexOf('alt="');
                return match.substring(index + 5, index + 7);
            }

            return '';
        });
    }

    return rocketChatMessage;
};

export const mapRocketChatMessageToTeamsMessage = (rocketChatMessage: string, originalSenderName?: string) : string => {
    // Handle emoji in text
    let teamsMessage = shortnameToUnicode(rocketChatMessage);

    const urlPattern = /(http:\/\/|https:\/\/)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)?/gi;
    const newLinePattern = /.*(\n)/g;

    teamsMessage = teamsMessage.replace(urlPattern, match => {
        return getTeamsMessageUrl(match);
    });

    teamsMessage = teamsMessage.replace(newLinePattern, match => {
        return `<p>${match}</p>`;
    });

    if (originalSenderName) {
        teamsMessage = getBridgedMessageFormat(originalSenderName, teamsMessage);
    }

    return teamsMessage;
};

const getTeamsMessageUrl = (url: string): string => {
    return `<a href=\"${url}\" title=\"${url}\" target=\"_blank\" rel=\"noreferrer noopener\">${url}</a>`;
};

const getBridgedMessageFormat = (originalSenderName: string, message: string): string => {
    return '<p style=\"font-size:14px; font-style:inherit; font-weight:inherit; margin-bottom:0; margin-left:0; margin-right:0; margin-top:0\">'
    + '<strong>[Bridged Message]</strong></p>'
    + '<blockquote style=\"font-size:14px; font-style:inherit; font-weight:inherit; margin:0.7rem 0\">'
    + '<p style=\"font-style:inherit; font-weight:inherit; margin-bottom:0; margin-left:0; margin-right:0; margin-top:0\">'
    + `<strong>${originalSenderName}:</strong></p>`
    + '<p style=\"font-style:inherit; font-weight:inherit; margin-bottom:0; margin-left:0; margin-right:0; margin-top:0\">'
    + message
    + '</p></blockquote>';
}
