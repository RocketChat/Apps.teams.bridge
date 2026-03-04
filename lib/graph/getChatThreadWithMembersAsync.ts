import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiChatThreadWithMemberUrl } from "../Const";
import { GetThreadResponse } from './types';
import { parseThreadType } from './parsers';

export const getChatThreadWithMembersAsync = async (
    http: IHttp,
    threadId: string,
    userAccessToken: string): Promise<GetThreadResponse> => {
    const url = getGraphApiChatThreadWithMemberUrl(threadId);

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get chat thread failed!');
        }

        let memberIds: string[] | undefined = undefined;

        const jsonMembers = responseBody.members as any[];
        if (jsonMembers) {
            memberIds = [];
            for (const jsonMember of jsonMembers) {
                memberIds.push(jsonMember.userId);
            }
        }

        const result: GetThreadResponse = {
            threadId: responseBody.id,
            topic: responseBody.topic,
            type: parseThreadType(responseBody.chatType),
            memberIds: memberIds,
        };

        return result;
    } else {
        throw new Error(`Get chat thread failed with http status code ${response.statusCode}.`);
    }
};
