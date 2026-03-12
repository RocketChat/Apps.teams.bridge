import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiUserUrl } from "../Const";
import { TeamsUserProfile } from './types';

export interface SearchTeamsUsersResult {
    users: TeamsUserProfile[];
    nextLink?: string;
}

export const searchTeamsUsersAsync = async (
    http: IHttp,
    appAccessToken: string,
    options: { query?: string; pageUrl?: string },
): Promise<SearchTeamsUsersResult> => {
    let url: string;
    if (options.pageUrl) {
        url = options.pageUrl;
    } else {
        url = `${getGraphApiUserUrl()}?$select=displayName,id,mail,givenName,surname&$top=25`;
        if (options.query) {
            url += `&$filter=startswith(displayName,'${encodeURIComponent(options.query)}')`;
        }
    }

    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${appAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode !== HttpStatusCode.OK) {
        throw new Error(`Search users failed with http status code ${response.statusCode}.`);
    }

    const responseBody = response.data;
    if (!responseBody) {
        throw new Error('Search users failed: empty response body');
    }

    const userList = (responseBody.value ?? []) as any[];
    const users: TeamsUserProfile[] = userList.map((user) => ({
        displayName: user.displayName ?? '',
        givenName: user.givenName ?? '',
        surname: user.surname ?? '',
        mail: user.mail ?? '',
        id: user.id,
    }));

    return {
        users,
        nextLink: responseBody['@odata.nextLink'],
    };
};
