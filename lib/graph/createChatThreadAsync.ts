import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiChatUrl } from "../Const";
import { CreateThreadResponse } from './types';

export const createChatThreadAsync = async (
    http: IHttp,
    membersTeamsIds: string[],
    roomName: string,
    userAccessToken: string): Promise<CreateThreadResponse> => {
    const url = getGraphApiChatUrl();

    const uniqueMembersTeamsIds = [...new Set(membersTeamsIds)];

    const members: any[] = [];
    for (const teamsIds of uniqueMembersTeamsIds) {
        const member = {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            'roles': ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${teamsIds}')`,
        };
        members.push(member);
    }

    const body = {
        'chatType': 'group',
        'members': members,
        'topic': roomName,
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
            throw new Error('Create group chat thread failed!');
        }

        const result: CreateThreadResponse = {
            threadId: responseBody.id,
        };

        return result;
    } else {
        throw new Error(
            `Create group chat thread failed with http status code ${
                response.statusCode
            }.\nReceived: ${JSON.stringify(response.data, null, 2)}`
        );
    }
};
