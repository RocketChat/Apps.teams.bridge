import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiUserUrl } from "../Const";
import { TeamsUserProfile } from './types';

export const listTeamsUserProfilesAsync = async (
    http: IHttp,
    appAccessToken: string): Promise<TeamsUserProfile[]> => {
    const url = getGraphApiUserUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${appAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('List users failed!');
        }

        const userList = responseBody.value as any[];
        const result: TeamsUserProfile[] = [];
        for (let index = 0; index < userList.length; index++) {
            try {
                const user = userList[index];
                const record: TeamsUserProfile = {
                    displayName: user.displayName,
                    givenName: user.givenName,
                    surname: user.surname,
                    mail: user.mail,
                    id: user.id,
                };
                result.push(record);
            } catch (error) {
                console.error(`Error when handling user list. Details: ${error}`);
            }
        }

        return result;
    } else {
        throw new Error(`List users failed with http status code ${response.statusCode}.`);
    }
};
