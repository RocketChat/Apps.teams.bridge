import { IAppAccessors } from "@rocket.chat/apps-engine/definition/accessors";
import { IApiEndpointMetadata } from "@rocket.chat/apps-engine/definition/api";
import { TestEnvironment } from "./Const";

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
