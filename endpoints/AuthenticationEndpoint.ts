import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    ApiEndpoint,
    IApiEndpointInfo,
    IApiRequest,
    IApiResponse,
} from "@rocket.chat/apps-engine/definition/api";

export class AuthenticationEndpoint extends ApiEndpoint {
    public path = 'auth';

    public async get(
        request: IApiRequest,
        endpoint: IApiEndpointInfo,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<IApiResponse>
    {
        // Now this is an empty endpoint for having an endpoint URL under App Info.
        // The organization admin need to config this URL in the AAD app they created for Teams interop.
        // TODO: implement this empoint.
        return this.success();
    }
}
