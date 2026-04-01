import {
    IHttp,
    IMessageBuilder,
    IModify,
    IModifyCreator,
    IPersistence,
    IRead,
    IRoomBuilder,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage, IMessageAttachment } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom, RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { IUploadDescriptor } from "@rocket.chat/apps-engine/definition/uploads/IUploadDescriptor";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { shortnameToUnicode } from "emojione";
import { TeamsAttachmentType } from "./Const";
import { downloadOneDriveFileAsync, getMessageAttachments, GetMessageResponse, MessageContentType } from "./MicrosoftGraphApi";
import { buildRocketChatMessageText, extractMainTextNodesFromBridgedMessageNodes, parseHTML } from "./TeamsMessageParser";
import { attachAttachments, attachMessageReferences, createTeamsHTMLMessage } from "./RocketChatMessageParser";
import { UploadMapping } from "./PersistHelper";
import type { MessageMappingModel, UploadMappingModel } from "./PersistHelper";
import { getAppAccessTokenAsync } from "./AuthHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";

export const sendRocketChatOneOnOneMessageAsync = async (
    message: string,
    sender: IUser,
    receiver: IUser,
    read: IRead,
    modify: IModify): Promise<string> => {
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


export const getAvatarUrlForUsername = async (username: string, read: IRead): Promise<string> => {
    const siteUrl = await read.getEnvironmentReader().getServerSettings().getValueById('Site_Url');
    return `${siteUrl || ''}/avatar/${username}`;
}

export const sendRocketChatMessageInRoomAsync = async (
    messageText: string,
    sender: IUser,
    room: IRoom,
    modify: IModify,
    read: IRead,
    options?: {
        alias?: string;
    }
): Promise<string> => {
    const creator: IModifyCreator = modify.getCreator();

    const message: IMessage = {
        text: messageText,
        sender,
        room,
        attachments: [await buildExtraInfoAttachment({ source: 'ms-teams', ...(options?.alias && { alias: options.alias }) })],
    };

    const messageBuilder: IMessageBuilder = creator.startMessage(message as IMessage);
    return await creator.finish(messageBuilder);
};

export const generateUploadCallback = ({
    attachments,
    uploadFiles,
    read,
    persistence,
    http,
    app,
    accessToken,
    room,
    sender,
    modify,
    alias,
}: {
    attachments: any[];
    uploadFiles: boolean;
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
    accessToken: string;
    room: IRoom;
    sender: IUser;
    modify: IModify;
    alias?: string;
}) => async (): Promise<{ rocketChat: string; teams: string }[]> => {
    let uploadIds: { rocketChat: string; teams: string }[] = [];
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
                                const appAccessToken = await getAppAccessTokenAsync(
                                    {
                                        read,
                                        persistence,
                                        http,
                                        app,
                                    },
                                );
                                const upload = await downloadAttachmentFileFromExternalAndUploadToRocketChatAsync(
                                    {
                                        url: contentUrl,
                                        fileName: name,
                                        accessToken: appAccessToken ?? accessToken,
                                        room,
                                        sender,
                                        http,
                                        modify,
                                        options: { alias },
                                    }
                                );
                                return { rocketChat: upload.id, teams: id };
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
            ).forEach((uploadData) => {
                if (uploadData) {
                    uploadIds.push(uploadData);
                }
            })
        }
    }
    return uploadIds;
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
    persistence,
    app,
    messageOptions,
}: {
    getMessageResponse: GetMessageResponse,
    read: IRead,
    accessToken: string
    room: IRoom,
    sender: IUser,
    modify: IModify,
    http: IHttp,
    uploadFiles: boolean,
    persistence: IPersistence,
    app: TeamsBridgeApp,
    messageOptions?: { alias?: string },
}): Promise<{
    text: string;
    uploadCallback: () => Promise<{
        rocketChat: string;
        teams: string;
    }[]>;
}> => {
    const { messageContent, messageContentType, attachments } = getMessageResponse;
    let text = messageContent;

    const uploadCallback = generateUploadCallback({
        attachments: attachments ?? [],
        uploadFiles,
        read,
        persistence,
        http,
        app,
        accessToken,
        room,
        sender,
        modify,
        alias: messageOptions?.alias,
    });

    if (messageContentType && messageContentType === MessageContentType.Html) {
        const isBridged = isBridgedMessageFormat(messageContent);
        const parsedNodes = parseHTML(messageContent);
        text = await buildRocketChatMessageText({ nodes: isBridged ? extractMainTextNodesFromBridgedMessageNodes(parsedNodes) : parsedNodes, attachments, read });
    }

    return {
        text,
        uploadCallback,
    }
};

export const mapRocketChatMessageToTeamsMessage = (rocketChatMessage: string, originalSenderName?: string): string => {
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
        teamsMessage = getBridgedMessageFormatV2(originalSenderName, teamsMessage);
    }

    return teamsMessage;
};

export const isBridgedMessageFormat = (message: string): boolean => {
    return message.includes('[Bridged Message]');
};

export const getBridgedMessageFormatV2 = (
    originalSenderName: string,
    message: string,
): string => {
    return (
        // Opening [Bridged Message] paragraph
        '<p style="font-size:14px; font-style:inherit; font-weight:inherit; margin-bottom:0; margin-left:0; margin-right:0; margin-top:0">' +
        "<strong>[Bridged Message]</strong>" +
        "</p>" +
        // Opening blockquote for ms-teams
        '<blockquote style="font-size:14px; font-style:inherit; font-weight:inherit; margin:0.7rem 0">' +
        // Sender name paragraph
        '<p style="font-style:inherit; font-weight:inherit; margin-bottom:0; margin-left:0; margin-right:0; margin-top:0">' +
        `<strong>${originalSenderName}:</strong>` +
        '<hr/>' +
        `</p>` +
        // Message paragraph
        '<p style="font-style:inherit; font-weight:inherit; margin-bottom:0; margin-left:0; margin-right:0; margin-top:0">' +
        message +
        '</p>' +
        // Closing blockquote
        '</blockquote>'
    );
};

export const mapRocketChatMessageToTeamsMessageV2 = async ({
    message,
    originalSenderName,
    read,
    forceBridgedMessage,
    siteUrl,
    http,
    accessToken,
    messageIdMapping,
}: {
    message: IMessage,
    originalSenderName?: string,
    read: IRead,
    forceBridgedMessage?: boolean,
    siteUrl?: string,
    http: IHttp,
    accessToken: string,
    messageIdMapping: MessageMappingModel,
}) => {
    // Handle emoji in text
    const text = message.text ?? "";
    const md = message[`_unmappedProperties_`]?.['md'] ?? [];
    if (md.length === 0 && text) {
        return {
            text: mapRocketChatMessageToTeamsMessage(text, originalSenderName),
            attachments: []
        };
    }

    const _siteUrl = siteUrl ?? await read.getEnvironmentReader().getServerSettings().getValueById("Site_Url") ?? "";
    let teamsMessage = createTeamsHTMLMessage(md, _siteUrl);
    const { html: teamsMessageWithReferences, attachments: messageAttachments } = await attachMessageReferences(read, http, teamsMessage, accessToken);
    const attachmentsFromTeams = await getMessageAttachments({
        http,
        messageId: messageIdMapping.teamsMessageId,
        threadId: messageIdMapping.teamsThreadId,
        userAccessToken: accessToken,
    });
    const attachmentsWithoutMessageAttachment = attachmentsFromTeams.filter(att => !messageAttachments.find(a => a?.id === att?.id));
    const finalAttachments = [...messageAttachments, ...attachmentsWithoutMessageAttachment];
    teamsMessage = attachAttachments({ html: teamsMessageWithReferences, attachmentIds: attachmentsWithoutMessageAttachment.map(a => a.id) });

    if (forceBridgedMessage) {
        teamsMessage = getBridgedMessageFormatV2(
            originalSenderName || 'Rocket.Chat User',
            teamsMessage,
        );
    }

    return { text: teamsMessage, attachments: finalAttachments };
};

const downloadAttachmentFileFromExternalAndUploadToRocketChatAsync = async ({
    url,
    fileName,
    accessToken,
    room,
    sender,
    http,
    modify,
    options,
}: {
    url: string,
    fileName: string,
    accessToken: string,
    room: IRoom,
    sender: IUser,
    http: IHttp,
    modify: IModify,
    options?: {
        alias?: string;
    }
}) => {
    const encodedUrl = `u!${base64Encode(url).replace(/=+$/, '').replace('/', '_').replace('+', '-')}`;

    const buff = await downloadOneDriveFileAsync(http, encodedUrl, accessToken);
    const uploadCreator = modify.getCreator().getUploadCreator();
    const fileInfo: IUploadDescriptor = {
        filename: buildExtraInfoFileName(fileName, { source: 'ms-teams', ...(options?.alias && { alias: options.alias }) }),
        room: room,
        user: sender
    };

    return uploadCreator.uploadBuffer(buff, fileInfo);
};

const getTeamsMessageUrl = (url: string): string => {
    return `<a href=\"${url}\" title=\"${url}\" target=\"_blank\" rel=\"noreferrer noopener\">${url}</a>`;
};

const base64Encode = (str: string): string => Buffer.from(str, 'binary').toString('base64');

export const buildExtraInfoAttachment = (data: any) => {
    const attachment: IMessageAttachment = {
        imageUrl: `data:image/svg+xml,<?xml version="1.0" encoding="UTF-8"?><svg viewBox="0 0 120 12" xmlns="http://www.w3.org/2000/svg"><text dominant-baseline="hanging" fill="currentColor" font-family="Arial, sans-serif" font-size="12">by MS Teams Bridge</text><metadata>${JSON.stringify({ extraInfo: data })}</metadata></svg>`,
        type: 'image/svg+xml',
        collapsed: true,
    };
    return attachment;
};

export const popExtraInfoAttachment = (message: IMessage) => {
    const data = {} as Record<string, any>;
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

    const [, baseName, encoded] = match;
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

export const combineRocketChatMessagesToTeamsMessage = async ({
    read,
    uploadMappings,
    messages,
    deletedMessages,
    deletedUploads,
    forceBridgedMessage,
    originalSenderName,
    messageIdMapping,
    http,
    accessToken,
}: {
    read: IRead;
    messageIdMapping: MessageMappingModel;
    messages: IMessage[];
    uploadMappings?: UploadMappingModel[];
    deletedMessages?: Set<string>;
    deletedUploads?: Set<string>;
    originalSenderName?: string;
    forceBridgedMessage?: boolean;
    http: IHttp;
    accessToken: string;
}) => {

    const targetMessages = messages.filter(
        (message) => message.id && !deletedMessages?.has(message.id)
    );

    const results = (
        await Promise.all(
            targetMessages.map((message) =>
                mapRocketChatMessageToTeamsMessageV2({
                    read,
                    message,
                    originalSenderName,
                    forceBridgedMessage,
                    http,
                    accessToken,
                    messageIdMapping,
                })
            )
        )
    );

    const text = results.map(r => r.text).join('\n');

    const parsedAttachments = results.map(r => r.attachments).reduce((acc, curr) => {
        acc.push(...curr);
        return acc;
    }, []);

    let targetUploadMappings: UploadMappingModel[];
    if (uploadMappings) {
        targetUploadMappings = uploadMappings.filter(
            (uploadMap) => !deletedUploads?.has(uploadMap.rocketchatUploadId)
        );
    } else {
        targetUploadMappings = (await UploadMapping.findByTeamsMessageId(
            read,
            messageIdMapping.teamsMessageId
        )).filter(
            (uploadMap) => !deletedUploads?.has(uploadMap.rocketchatUploadId)
        );
    }

    const teamsAttachmentIds = targetUploadMappings.map(um => um.teamsAttachmentId);
    const attachmentsFromTeams = await getMessageAttachments({
        http,
        messageId: messageIdMapping.teamsMessageId,
        threadId: messageIdMapping.teamsThreadId,
        userAccessToken: accessToken,
    })

    // Remove duplicate attachments by id
    const allAttachments = [...parsedAttachments, ...attachmentsFromTeams];
    const uniqueAttachmentsMap = new Map<string, any>();
    for (const att of allAttachments) {
        if (att && att.id && !uniqueAttachmentsMap.has(att.id)) {
            uniqueAttachmentsMap.set(att.id, att);
        }
    }
    const uniqueAttachments = Array.from(uniqueAttachmentsMap.values());
    const finalAttachments = filterTeamsAttachments(uniqueAttachments, teamsAttachmentIds);

    return {
        text: attachAttachments({ html: text, attachmentIds: teamsAttachmentIds }),
        shouldDeleteTeamsMessage:
            targetMessages.length === 0 && teamsAttachmentIds.length === 0,
        attachments: finalAttachments,
    };
}

export const filterTeamsAttachments = (attachments: any[], toKeepIds: string[]): any[] => {
    return attachments.filter(att => toKeepIds.includes(att.id));
};
