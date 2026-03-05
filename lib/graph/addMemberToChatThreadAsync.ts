import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiChatMemberUrl } from "../Const";

export const addMemberToChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    memberTeamsIdToBeAdd: string,
    userAccessToken: string): Promise<void> => {
    const url = getGraphApiChatMemberUrl(threadId);

    const body = {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'roles': ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${memberTeamsIdToBeAdd}')`,
        'visibleHistoryStartDateTime': '0001-01-01T00:00:00Z',
    };

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body)
    };

    const response = await http.post(url, httpRequest);
    console.log(JSON.stringify(response, null, 2));

    if (response.statusCode !== HttpStatusCode.CREATED) {
        throw new Error(`Add member to group chat thread failed with http status code ${response.statusCode}.`);
    }
};
