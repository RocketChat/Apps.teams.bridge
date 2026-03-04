import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiChatMemberRemoveUrl } from "../Const";

export const removeMemberFromChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    teamsUserId: string,
    userAccessToken: string): Promise<void> => {
    const url = getGraphApiChatMemberRemoveUrl(threadId, teamsUserId);
    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.del(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        return;
    } else {
        throw new Error(`Remove member from chat thread failed with http status code ${response.statusCode}.`);
    }
};
