import {
    HttpStatusCode,
    IHttp,
    IHttpRequest,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { RocketChatAssociationModel, RocketChatAssociationRecord } from "@rocket.chat/apps-engine/definition/metadata";
import { AppSetting } from "../config/Settings";

export const getApplicationAccessTokenAsync = async (read: IRead, http: IHttp, persis: IPersistence) : Promise<boolean> => {
    const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
    const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
    const aadClientSecret = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientSecret)).value;

    const requestBody = "scope=https://graph.microsoft.com/.default&grant_type=client_credentials"
        + `&client_id=${aadClientId}&client_secret=${aadClientSecret}`;

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/x-www-form-urlencoded"
        },
        content: requestBody,
    };

    const url = `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/token`;
    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            return false;
        }

        const jsonBody = JSON.parse(responseBody);

        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, 'ApplicationAccessToken'),
        ];
        const data = {
            'ApplicationAccessToken': jsonBody.access_token,
        };
        await persis.updateByAssociations(associations, data, true);

        return true;
    } else {
        return false;
    }
};
