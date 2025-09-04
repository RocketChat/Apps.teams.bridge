import {
    IHttp,
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
import { LoginButtonText, MicrosoftFileUrlPrefix, SharePointUrl, TeamsAttachmentType } from "./Const";
import { downloadOneDriveFileAsync, GetMessageResponse, MessageContentType } from "./MicrosoftGraphApi";
import { buildRocketChatMessageText, parseHTML } from "./TeamsMessageParser";

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
        room,
        attachments: [await buildExtraInfoAttachment({ source: 'ms-teams' })],
    }

    const messageBuilder: IMessageBuilder = creator.startMessage(message as IMessage);
    return creator.finish(messageBuilder);
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

export const mapTeamsMessageToRocketChatMessage = async ({
    getMessageResponse,
    read,
    accessToken,
    modify,
    room,
    sender,
    http,
    uploadFiles,
}: {
    getMessageResponse: GetMessageResponse,
    read: IRead,
    accessToken: string
    room: IRoom,
    sender: IUser,
    modify: IModify,
    http: IHttp,
    uploadFiles: boolean,
}): Promise<{
    text: string;
    uploadIds: {
        rocketChat: string;
        teams: string;
    }[];
}> => {
    const { messageContent, messageContentType, attachments } = getMessageResponse;
    let uploadIds: { rocketChat: string; teams: string }[] = [];
    let text = messageContent;
    if (attachments && attachments.length > 0) {
        if (uploadFiles) {
            const attachmentsToUpload = attachments.filter(
                (a) => a.contentType === TeamsAttachmentType.File
            );
           (
                await Promise.all(
                    attachmentsToUpload.map(async (attachment) => {
                        const { contentUrl, name, id } = attachment;
                        if (contentUrl && name) {
                            try {
                                const upload = await downloadAttachmentFileFromExternalAndUploadToRocketChatAsync(
                                    {
                                        url: contentUrl,
                                        fileName: name,
                                        accessToken,
                                        room,
                                        sender,
                                        http,
                                        modify,
                                    }
                                );
                                return upload;
                            } catch (error) {
                                console.error(
                                    `Error downloading attachment: ${error.message}`
                                );
                                return Promise.resolve(null);
                            }
                        }
                        return Promise.resolve(null);
                    })
                )
            ).forEach((upload, index) => {
                if (upload) {
                    uploadIds.push({ rocketChat: upload.id, teams: attachments[index].id });
                }
            })
        }
    }
    if (messageContentType && messageContentType === MessageContentType.Html) {
        const parsedNodes = parseHTML(messageContent);
        text = await buildRocketChatMessageText({ nodes: parsedNodes, attachments, read });
    }

    return {
        text,
        uploadIds,
    }
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

const downloadAttachmentFileFromExternalAndUploadToRocketChatAsync = async ({
    url,
    fileName,
    accessToken,
    room,
    sender,
    http,
    modify,
}: {
    url: string,
    fileName: string,
    accessToken: string,
    room: IRoom,
    sender: IUser,
    http: IHttp,
    modify: IModify,
}) => {
    const encodedUrl = `u!${base64Encode(url).replace(/=+$/, '').replace('/', '_').replace('+', '-')}`;

    const buff = await downloadOneDriveFileAsync(http, encodedUrl, accessToken);
    const uploadCreator = modify.getCreator().getUploadCreator();
    const fileInfo: IUploadDescriptor = {
        filename: buildExtraInfoFileName(fileName, { source: 'ms-teams' }),
        room: room,
        user: sender
    };

    return uploadCreator.uploadBuffer(buff, fileInfo);
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

export const buildExtraInfoAttachment = (data: any) => {
    const attachment: IMessageAttachment = {
        imageUrl: `data:image/svg+xml,<?xml version="1.0" encoding="UTF-8"?><svg viewBox="0 0 120 12" xmlns="http://www.w3.org/2000/svg"><text dominant-baseline="hanging" fill="currentColor" font-family="Arial, sans-serif" font-size="12">by MS Teams Bridge</text><metadata>${JSON.stringify({ extraInfo: data })}</metadata></svg>`,
        type: 'image/svg+xml',
        collapsed: true,
    };
    return attachment;
};

export const popExtraInfoAttachment = (message: IMessage) => {
    const data = {} as any;
    const extraInfoAttachment = message.attachments?.find((att) => {
        const match = att.imageUrl?.match(/<metadata>([\s\S]*?)<\/metadata>/);
        if (match) {
            try {
                const parsedData = JSON.parse(match[1]);
                if ('extraInfo' in parsedData) {
                    Object.assign(data, parsedData.extraInfo);
                    return true;
                }
                return false;
            } catch {
                return false;
            }
        }
        return false;
    });
    if (extraInfoAttachment) {
        message.attachments = message.attachments?.filter(att => att !== extraInfoAttachment);
    }
    return data;
};


export const buildExtraInfoFileName = (
    originalFilename: string,
    data: any
): string => {
    // Encode metadata as base64
    const encoded = Buffer.from(
        JSON.stringify({ extraInfo: data }),
        "utf8"
    ).toString("base64");

    // Insert before extension
    const lastDot = originalFilename.lastIndexOf(".");
    if (lastDot !== -1) {
        const name = originalFilename.slice(0, lastDot);
        const ext = originalFilename.slice(lastDot); // includes the dot
        return `${name}__extradata_${encoded}${ext}`;
    }

    // No extension case
    return `${originalFilename}__extradata_${encoded}`;
};
export const getExtraInfoAndOriginalFileName = (filename: string): { originalFilename: string; extraInfo: any; present: boolean } => {
    if (!filename) {
        return { originalFilename: '', extraInfo: {}, present: false };
    }

    // Match: <name>__extradata_<base64>[.<ext>]
    const match = filename.match(/^(.*)__extradata_([^.]*)/);
    if (!match) {
        return { originalFilename: filename, extraInfo: {}, present: false };
    }

    const [ , baseName, encoded ] = match;
    const ext = filename.slice((baseName + `__extradata_${encoded}`).length);

    try {
        const decoded = Buffer.from(encoded, "base64").toString("utf8");
        const { extraInfo } = JSON.parse(decoded);
        return {
            originalFilename: baseName + ext,
            extraInfo: extraInfo ?? {},
            present: true,
        };
    } catch {
        return { originalFilename: filename, extraInfo: {}, present: false };
    }
};
