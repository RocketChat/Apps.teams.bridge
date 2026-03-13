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
import { LoginMessage, UserMapping, UserRegistration, AppUserLoginNotified, OAuthNonce } from "../lib/PersistHelper";
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
            this.app.getLogger().warn(
                `Authentication failed — AAD error: ${request.query.error}` +
                (request.query.error_description ? ` — ${request.query.error_description}` : ``)
            );
            return this.errorResponse();
        }

        let rocketChatUserId: string;
        let type: 'bot' | 'normal';
        let aadTenantId: string;
        let aadClientId: string;
        let aadClientSecret: string;
        let accessCode: string;
        let authEndpointUrl: string;
        try {
            const parsed = JSON.parse(
                Buffer.from(request.query.state, "base64").toString("utf-8"),
            ) as Record<string, unknown>;
            if (
                typeof parsed.rc_uid !== 'string' || parsed.rc_uid.length === 0 ||
                (parsed.type !== 'bot' && parsed.type !== 'normal')
            ) {
                this.app.getLogger().warn('Authentication rejected — malformed state parameter.');
                return this.errorResponse();
            }
            rocketChatUserId = parsed.rc_uid;
            type = parsed.type as 'bot' | 'normal';

            if (typeof parsed.nonce !== 'string' || parsed.nonce.length === 0) {
                this.app.getLogger().warn('Authentication rejected — no nonce in state.');
                return this.errorResponse();
            }
            const storedNonce = await OAuthNonce.findAndDelete(read, persis, rocketChatUserId);
            if (storedNonce === null || storedNonce !== parsed.nonce) {
                this.app.getLogger().warn('Authentication rejected — nonce mismatch.');
                return this.errorResponse();
            }

            aadTenantId = (
                await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)
            ).value;
            aadClientId = (
                await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)
            ).value;
            aadClientSecret = (
                await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientSecret)
            ).value;
            accessCode = request.query.code;
            authEndpointUrl = await getRocketChatAppEndpointUrl(
                this.app.getAccessors(),
                AuthenticationEndpointPath
            );
        } catch (error) {
            this.app.getLogger().error(
                `Authentication — phase A (state/settings) failed: ${(error as Error)?.message ?? String(error)}`
            );
            return this.errorResponse();
        }

        let userAccessToken: string;
        let refreshToken: string;
        let expiresIn: number;
        let extExpiresIn: number;
        let teamsUserId: string;
        let appUser: Awaited<ReturnType<typeof read.getUserReader.prototype.getAppUser>>;
        try {
            const response = await getUserAccessTokenAsync(
                http,
                accessCode,
                authEndpointUrl,
                aadTenantId,
                aadClientId,
                aadClientSecret,
                type,
            );

            if (!response.refreshToken) {
                throw new Error('Token exchange did not return a refresh token.');
            }

            userAccessToken = response.accessToken;
            refreshToken = response.refreshToken;
            expiresIn = response.expiresIn;
            extExpiresIn = response.extExpiresIn;

            const teamsUserProfile = await getUserProfileAsync(http, userAccessToken);
            teamsUserId = teamsUserProfile.id;

            appUser = await read.getUserReader().getAppUser(this.app.getID());

            const existingMapping = await UserMapping.findByTeamsUserId(read, teamsUserId);
            if (existingMapping !== null && existingMapping.rocketChatUserId !== rocketChatUserId) {
                if (appUser && existingMapping.rocketChatUserId === appUser.id) {
                    return this.conflictResponse(
                        "This Teams account is used by the app user and cannot be linked to a personal Rocket.Chat account.",
                    );
                }
                return this.conflictResponse(
                    "This Teams account is already linked to another Rocket.Chat user. Please log out from Teams on the other account first.",
                );
            }
        } catch (error) {
            this.app.getLogger().error(
                `Authentication — phase B (token exchange/profile) failed: ${(error as Error)?.message ?? String(error)}`
            );
            return this.errorResponse();
        }

        try {
            await Promise.all([
                UserRegistration.persist(
                    persis,
                    rocketChatUserId,
                    userAccessToken,
                    refreshToken,
                    expiresIn,
                    extExpiresIn
                ),
                UserMapping.persist(persis, rocketChatUserId, teamsUserId),
                LoginMessage.save({
                    persistence: persis,
                    rocketChatUserId,
                    wasSent: false,
                }),
            ]);

            // If the app user just logged in, reset the per-room notification
            // flags so rooms won't show stale "app user not logged in" warnings.
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
                teamsUserId,
                userAccessToken,
                renewIfExists: true,
                forceRenew: true,
            });

            return this.success(this.embeddedLoginSuccessMessage);
        } catch (error) {
            this.app.getLogger().error(
                `Authentication — phase C (persistence/subscription) failed: ${(error as Error)?.message ?? String(error)}`
            );
            return this.errorResponse();
        }
    }

    private conflictResponse(message: string): IApiResponse {
        const response: IApiResponseJSON = {
            status: HttpStatusCode.CONFLICT,
            content: { message },
        };
        return response;
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
