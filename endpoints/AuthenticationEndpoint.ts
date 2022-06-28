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
import { AuthenticationEndpointPath } from "../lib/TeamsBridgeConst";

export class AuthenticationEndpoint extends ApiEndpoint {
    private embeddedLoginSuccessMessage: string = 'Login to Teams succeed! You can close this window now.'

    public path = AuthenticationEndpointPath;

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

        // 1. Get RC user id
        // 2. Get user access Token & fresh Token
        // 3. Persist the access
        // 4. Setup token refresh mechenism

        // TODO: setup incoming message webhook with Microsoft

        return this.success(this.embeddedLoginSuccessMessage);
    }
}
