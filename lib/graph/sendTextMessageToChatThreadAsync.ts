import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiMessageUrl } from "../Const";
import { SendMessageResponse } from './types';

export const sendTextMessageToChatThreadAsync = async ({
    http,
    textMessage,
    threadId,
    userAccessToken,
    attachments,
}: {
    http: IHttp;
    textMessage: string;
    threadId: string;
    userAccessToken: string;
    attachments?: any[];
}): Promise<SendMessageResponse> => {
    const url = getGraphApiMessageUrl(threadId);

    const body = {
        'body': {
            'content': textMessage,
            'contentType': 'html'
        },
        ...(attachments && { attachments })
    };

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body)
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Send message to chat thread failed!');
        }

        const result: SendMessageResponse = {
            messageId: responseBody.id,
        };

        return result;
    } else {
        throw new Error(`Send message to chat thread failed with http status code ${response.statusCode}.`);
    }
};
