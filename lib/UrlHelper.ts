import { IAppAccessors } from "@rocket.chat/apps-engine/definition/accessors";
import { IApiEndpointMetadata } from "@rocket.chat/apps-engine/definition/api";

// Put url here when running locally & using tunnel service such as Ngrok to expose the localhost port to the internet
const tunnelServiceUrl: string = '';

export const getRocketChatAppEndpointUrl = async (appAccessors: IAppAccessors, appEndpointPath: string) : Promise<string> => {

    const webhookEndpoint: IApiEndpointMetadata = appAccessors.providedApiEndpoints
        .find((endpoint) => endpoint.path === appEndpointPath) as IApiEndpointMetadata;
    let siteUrl: string = await appAccessors.environmentReader.getServerSettings().getValueById('Site_Url');
    siteUrl = siteUrl.substring(0, siteUrl.length - 1);
    
    if (tunnelServiceUrl && tunnelServiceUrl !== '') {
        siteUrl = tunnelServiceUrl;
    }

    return siteUrl + webhookEndpoint.computedPath;
};
