import {
    IHttp,
    IHttpRequest,
    IMessageBuilder,
    IModify,
    IModifyCreator,
    INotifier,
    IRead,
    IRoomBuilder,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage, IMessageAction, IMessageAttachment, MessageActionType } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom, RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUploadDescriptor } from "@rocket.chat/apps-engine/definition/uploads/IUploadDescriptor";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { shortnameToUnicode } from "emojione";
import { FileAttachmentContentType, LoginButtonText, MicrosoftFileUrlPrefix, SharePointUrl } from "./Const";
import { downloadOneDriveFileAsync, GetMessageResponse, MessageContentType } from "./MicrosoftGraphApi";

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
    messageText: string,
    sender: IUser,
    room: IRoom,
    modify: IModify) : Promise<string> => {
    const creator: IModifyCreator = modify.getCreator();

    const message : IMessage = {
        text: messageText,
        sender,
        room
    }

    const messageBuilder: IMessageBuilder = creator.startMessage(message as IMessage);
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
    getMessageResponse: GetMessageResponse,
    accessToken: string,
    room: IRoom,
    sender: IUser,
    http: IHttp,
    modify: IModify) : string => {
    const teamsMessage = getMessageResponse.messageContent;
    const contentType = getMessageResponse.messageContentType;

    console.log("Mapping Teams message format to Rocket.Chat message format.");

    let rocketChatMessage = teamsMessage;
    if (contentType && contentType === MessageContentType.Html) {
        // TODO: find a better way to trim html tag from html messages
        const nbspPattern = /&nbsp;/g;
        const htmpTagPattern = /<\/?[^>]+>/g;

        console.log("Processing Teams Message!");

        rocketChatMessage = rocketChatMessage.replace(nbspPattern, ' ');
        rocketChatMessage = rocketChatMessage.replace(htmpTagPattern, match => {
            if (match.indexOf('itemtype="http://schema.skype.com/Emoji"') > 0) {
                console.log("find emoji!");
                const index = match.indexOf('alt="');
                return match.substring(index + 5, index + 7);
            }

            if (match.indexOf('img') > 0) {
                console.log("find img!");
                const urlStartIndex = match.indexOf('src="');
                const urlPrefixString = match.substring(urlStartIndex + 5);
                const urlEndIndex = urlPrefixString.indexOf('"');
                const url = urlPrefixString.substring(0, urlEndIndex);

                const fileNameStartIndex = match.indexOf('alt="');
                const fileNamePrefixString = match.substring(fileNameStartIndex + 5);
                const fileNameEndIndex = fileNamePrefixString.indexOf('"');
                const fileName = fileNamePrefixString.substring(0, fileNameEndIndex);

                console.log(`Find an URL: ${url}`);

                downloadInlineImgFromExternalAndUploadToRocketChatAsync(url, fileName, accessToken, room, sender, http, modify);

                return '';
            }

            if (match.indexOf('attachment') > 0) {
                console.log("find attachment!");
                const attachmentIdStartIndex = match.indexOf('id="');
                const attachmentIdPrefixString = match.substring(attachmentIdStartIndex + 4);
                const attachmentIdEndIndex = attachmentIdPrefixString.indexOf('"');
                const attachmentId = attachmentIdPrefixString.substring(0, attachmentIdEndIndex);

                if (getMessageResponse.attachments) {
                    for (const attachment of getMessageResponse.attachments) {
                        if (attachment.id === attachmentId && attachment.contentType === FileAttachmentContentType) {
                            const fileName = attachment.name;
                            const url = attachment.contentUrl;

                            console.log(`Find an URL: ${url}`);

                            downloadAttachmentFileFromExternalAndUploadToRocketChatAsync(url, fileName, accessToken, room, sender, http, modify);
                           
                            return '';
                        }
                    }
                }
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

const downloadAttachmentFileFromExternalAndUploadToRocketChatAsync = async (
    url: string,
    fileName: string,
    accessToken: string,
    room: IRoom,
    sender: IUser,
    http: IHttp,
    modify: IModify) : Promise<void> => {
    const encodedUrl = `u!${base64Encode(url).replace(/=+$/, '').replace('/','_').replace('+','-')}`;
    console.log(encodedUrl);

    const buff = await downloadOneDriveFileAsync(http, encodedUrl, accessToken);
    const uploadCreator = modify.getCreator().getUploadCreator();
    const fileInfo: IUploadDescriptor = {
        filename: fileName,
        room: room,
        user: sender
    };

    await uploadCreator.uploadBuffer(buff, fileInfo);
};

const downloadInlineImgFromExternalAndUploadToRocketChatAsync = async (
    url: string,
    fileName: string,
    accessToken: string,
    room: IRoom,
    sender: IUser,
    http: IHttp,
    modify: IModify) : Promise<void> => {
    const httpRequest: IHttpRequest = {
        encoding: null
    };

    if (url.startsWith(MicrosoftFileUrlPrefix) || url.indexOf(SharePointUrl)) {
        // Auth Required
        httpRequest.headers = {
            'Authorization': `Bearer ${accessToken}`,
        };
    }
    
    const response = await http.get(url, httpRequest);
    let fileMIMEType = '';
    if (response.headers) {
        fileMIMEType = response.headers['content-type'];
        console.log(response.headers['content-type']);
    }
    const imgStr = response.content as string;
    const buff = Buffer.from(imgStr, 'binary');
    
    const uploadCreator = modify.getCreator().getUploadCreator();
    const imgInfo: IUploadDescriptor = {
        filename: `${fileName}.${fileMIMEType.split('/')[1]}`,
        room: room,
        user: sender
    };

    await uploadCreator.uploadBuffer(buff, imgInfo);
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

const base64Encode = (str: string):string => Buffer.from(str, 'binary').toString('base64');

