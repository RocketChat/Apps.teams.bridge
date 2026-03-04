import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiShareUrl } from "../Const";

export const downloadOneDriveFileAsync = async (
    http: IHttp,
    encodedUrl: string,
    userAccessToken: string): Promise<Buffer> => {
    const url = getGraphApiShareUrl(encodedUrl);

    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
        },
        encoding: null
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const fileStr = response.content as string;
        const buff = Buffer.from(fileStr, 'binary');
        return buff;
    } else {
        throw new Error(`Download one drive file failed with http status code ${response.statusCode}.`);
    }
};
