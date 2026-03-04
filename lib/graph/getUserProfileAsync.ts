import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiProfileUrl } from "../Const";
import { TeamsUserProfile } from './types';

export const getUserProfileAsync = async (http: IHttp, userAccessToken: string): Promise<TeamsUserProfile> => {
    const url = getGraphApiProfileUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get user profile failed!');
        }

        const result: TeamsUserProfile = {
            displayName: responseBody.displayName,
            givenName: responseBody.givenName,
            surname: responseBody.surname,
            mail: responseBody.mail,
            id: responseBody.id,
        };

        return result;
    } else {
        throw new Error(`Get user profile failed with http status code ${response.statusCode}.`);
    }
};
