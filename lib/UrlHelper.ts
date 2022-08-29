import { IAppAccessors } from "@rocket.chat/apps-engine/definition/accessors";
import { IApiEndpointMetadata } from "@rocket.chat/apps-engine/definition/api";
import { AuthenticationScopes, getMicrosoftAuthorizeUrl, TestEnvironment } from "./Const";

export const getRocketChatAppEndpointUrl = async (appAccessors: IAppAccessors, appEndpointPath: string) : Promise<string> => {

    const webhookEndpoint: IApiEndpointMetadata = appAccessors.providedApiEndpoints
        .find((endpoint) => endpoint.path === appEndpointPath) as IApiEndpointMetadata;
    let siteUrl: string = await appAccessors.environmentReader.getServerSettings().getValueById('Site_Url');
    siteUrl = siteUrl.substring(0, siteUrl.length - 1);
    
    if (TestEnvironment.tunnelServiceUrl && TestEnvironment.tunnelServiceUrl !== '') {
        siteUrl = TestEnvironment.tunnelServiceUrl;
    }

    return siteUrl + webhookEndpoint.computedPath;
};

export const getLoginUrl = (
    aadTenantId: string,
    aadClientId: string,
    authEndpointUrl: string,
    userId: string): string => {
    let url = getMicrosoftAuthorizeUrl(aadTenantId);
    url += `?client_id=${aadClientId}`;
    url += '&response_type=code';
    url += `&redirect_uri=${authEndpointUrl}`;
    url += '&response_mode=query';
    url += `&scope=${AuthenticationScopes.join('%20')}`;
    url += `&state=${userId}`;

    return url;
};
