import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiShareOneDriveFileUrl } from "../Const";
import { ShareOneDriveFileResponse } from './types';

export const shareOneDriveFileAsync = async (
    http: IHttp,
    oneDriveItemId: string,
    userAccessToken: string): Promise<ShareOneDriveFileResponse> => {
    const url = getGraphApiShareOneDriveFileUrl(oneDriveItemId);

    const body = {
        'type': 'view',
        'scope': 'organization',
    };

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if ([HttpStatusCode.CREATED, HttpStatusCode.OK].includes(response.statusCode)) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Create share link for onedrive item failed!');
        }

        const result: ShareOneDriveFileResponse = {
            shareId: responseBody.id,
            shareLink: responseBody.link.webUrl,
        };

        return result;
    } else {
        throw new Error(`Create share link for onedrive item failed with http status code ${response.statusCode}.`);
    }
};
