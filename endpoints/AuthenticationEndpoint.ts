import {
    HttpStatusCode,
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
import { IApiResponseJSON } from "@rocket.chat/apps-engine/definition/api/IResponse";
import { IApp } from "@rocket.chat/apps-engine/definition/IApp";
import { AppSetting } from "../config/Settings";
import {
    AuthenticationEndpointPath,
    SubscriberEndpointPath,
} from "../lib/Const";
import {
    getUserAccessTokenAsync,
    getUserProfileAsync,
    subscribeToAllMessagesForOneUserAsync,
} from "../lib/MicrosoftGraphApi";
import { LoginMessage, UserMapping, UserRegistration, AppUserLoginNotified } from "../lib/PersistHelper";
import { getRocketChatAppEndpointUrl } from "../lib/UrlHelper";

export class AuthenticationEndpoint extends ApiEndpoint {
    private embeddedLoginSuccessMessage: string =
        "Login to Teams succeed! You can close this window now.";
    private embeddedLoginFailureMessage: string =
        "Login to Teams failed! Please check document or contact your organization admin.";

    public path = AuthenticationEndpointPath;

    constructor(app: IApp) {
        super(app);
        this.errorResponse = this.errorResponse.bind(this);
    }

    public async get(
        request: IApiRequest,
        endpoint: IApiEndpointInfo,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence
    ): Promise<IApiResponse> {
        if (request.query.error && !request.query.code) {
            return this.errorResponse();
        }

        try {
            const aadTenantId = (
                await read
                    .getEnvironmentReader()
                    .getSettings()
                    .getById(AppSetting.AadTenantId)
            ).value;
            const aadClientId = (
                await read
                    .getEnvironmentReader()
                    .getSettings()
                    .getById(AppSetting.AadClientId)
            ).value;
            const aadClientSecret = (
                await read
                    .getEnvironmentReader()
                    .getSettings()
                    .getById(AppSetting.AadClientSecret)
            ).value;

            const { rc_uid: rocketChatUserId, type } = JSON.parse(
                Buffer.from(request.query.state, "base64").toString("utf-8"),
            ) as {
                rc_uid: string;
                type: "bot" | "normal";
            };
            const accessCode: string = request.query.code;
            const authEndpointUrl = await getRocketChatAppEndpointUrl(
                this.app.getAccessors(),
                AuthenticationEndpointPath
            );

            const response = await getUserAccessTokenAsync(
                http,
                accessCode,
                authEndpointUrl,
                aadTenantId,
                aadClientId,
                aadClientSecret,
                type,
            );

            const userAccessToken = response.accessToken;

            const teamsUserProfile = await getUserProfileAsync(
                http,
                userAccessToken
            );

            await Promise.all([
                UserRegistration.persist(
                    persis,
                    rocketChatUserId,
                    userAccessToken,
                    response.refreshToken as string,
                    response.expiresIn,
                    response.extExpiresIn
                ),
                UserMapping.persist(persis, rocketChatUserId, teamsUserProfile.id),
                LoginMessage.save({
                    persistence: persis,
                    rocketChatUserId,
                    wasSent: false,
                }),
            ]);

            // If the app user just logged in, reset the per-room notification
            // flags so rooms won't show stale "app user not logged in" warnings.
            const appUser = await read.getUserReader().getAppUser(this.app.getID());
            if (appUser && rocketChatUserId === appUser.id) {
                await AppUserLoginNotified.clearAll(persis);
            }

            const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(
                this.app.getAccessors(),
                SubscriberEndpointPath
            );

            await subscribeToAllMessagesForOneUserAsync({
                http,
                read,
                persis,
                rocketChatUserId,
                subscriberEndpointUrl,
                teamsUserId: teamsUserProfile.id,
                userAccessToken,
                renewIfExists: true,
                forceRenew: true,
            });

            return this.success(this.embeddedLoginSuccessMessage);
        } catch (error) {
            console.log("Error in authentication endpoint:" + JSON.stringify(error, null, 2));
            return this.errorResponse();
        }
    }

    private errorResponse(): IApiResponse {
        const response: IApiResponseJSON = {
            status: HttpStatusCode.BAD_REQUEST,
            content: {
                message: this.embeddedLoginFailureMessage,
            },
        };

        return response;
    }
}
