import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiResourceUrl } from "../Const";
import { Attachment, GetMessageResponse } from './types';
import { parseMessageType, parseMessageContentType } from './parsers';
import { MessageType } from './types';

export const getMessageWithResourceStringAsync = async (
    http: IHttp,
    resourceString: string,
    userAccessToken: string): Promise<GetMessageResponse> => {
    const url = getGraphApiResourceUrl(resourceString);

    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get message with resource string failed!');
        }

        let attachments: Attachment[] | undefined = undefined;

        const jsonAttachments = responseBody.attachments as any[];
        if (jsonAttachments && jsonAttachments.length > 0) {
            attachments = [];
            for (const jsonAttachment of jsonAttachments) {
                const attachment: Attachment = {
                    id: jsonAttachment.id,
                    contentType: jsonAttachment.contentType,
                    contentUrl: jsonAttachment.contentUrl,
                    name: jsonAttachment.name,
                };
                attachments.push(attachment);
            }
        }

        const messageType = parseMessageType(responseBody.messageType, responseBody.eventDetail);

        let memberIds: string[] | undefined = undefined;
        if (messageType === MessageType.SystemAddMembers || messageType === MessageType.SystemRemoveMembers) {
            memberIds = [];
            const jsonMembers = responseBody.eventDetail.members as any[];
            for (const jsonMember of jsonMembers) {
                memberIds.push(jsonMember.id);
            }
        }

        const result: GetMessageResponse = {
            threadId: responseBody.chatId,
            messageId: responseBody.id,
            messageType: messageType,
            fromTeamsUser: {
                id: responseBody.from?.user?.id,
                displayName: responseBody.from?.user?.displayName,
            },
            messageContentType: parseMessageContentType(responseBody.body?.contentType),
            messageContent: responseBody.body?.content,
            attachments: attachments,
            memberIds: memberIds,
            reactions: responseBody.reactions,
        };

        return result;
    } else {
        throw new Error(`Get message with resource string failed with http status code ${response.statusCode}.`);
    }
};
