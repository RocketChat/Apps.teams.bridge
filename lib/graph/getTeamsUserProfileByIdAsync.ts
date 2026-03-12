import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiUserUrl } from "../Const";

export const getTeamsUserProfileByIdAsync = async (
    http: IHttp,
    appAccessToken: string,
    teamsUserId: string,
): Promise<{ displayName: string; mail: string } | null> => {
    const url = `${getGraphApiUserUrl()}/${teamsUserId}?$select=displayName,mail`;
    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${appAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode !== HttpStatusCode.OK) {
        return null;
    }

    const data = response.data;
    if (!data) {
        return null;
    }

    return {
        displayName: data.displayName ?? '',
        mail: data.mail ?? '',
    };
};
