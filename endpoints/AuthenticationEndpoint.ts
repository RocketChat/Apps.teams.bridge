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

        const stateResult = await this.validateStateAndNonce(request, read, persis);
        if (!stateResult) return this.errorResponse();
        const { rocketChatUserId, type, accessCode } = stateResult;

        const env = await this.fetchAppEnvironment(read);
        if (!env) return this.errorResponse();
        const { aadTenantId, aadClientId, aadClientSecret, authEndpointUrl } = env;

        const tokenResult = await this.exchangeAndValidateTokens(
            http, read, persis, accessCode, authEndpointUrl, aadTenantId, aadClientId, aadClientSecret, type, rocketChatUserId
        );
        if (!tokenResult) {
            return this.errorResponse();
        }
        if ('conflictError' in tokenResult) {
            return this.conflictResponse(tokenResult.conflictError as string);
        }

        const persistResult = await this.persistAndSubscribe(
            read, persis, http, rocketChatUserId, type, tokenResult
        );
        if (!persistResult) {
            return this.errorResponse();
        }

        return this.success(this.embeddedLoginSuccessMessage);
    }

    private async validateStateAndNonce(
        request: IApiRequest,
        read: IRead,
        persis: IPersistence
    ): Promise<{ rocketChatUserId: string; type: 'bot' | 'normal'; accessCode: string } | null> {
        try {
            const parsed = JSON.parse(
                Buffer.from(request.query.state, "base64").toString("utf-8"),
            ) as Record<string, unknown>;
            if (
                typeof parsed.rc_uid !== 'string' || parsed.rc_uid.length === 0 ||
                (parsed.type !== 'bot' && parsed.type !== 'normal')
            ) {
                this.app.getLogger().warn('Authentication rejected — malformed state parameter.');
                return null;
            }

            const rocketChatUserId = parsed.rc_uid as string;
            const type = parsed.type as 'bot' | 'normal';

            if (typeof parsed.nonce !== 'string' || parsed.nonce.length === 0) {
                this.app.getLogger().warn('Authentication rejected — no nonce in state.');
                return null;
            }
            const storedNonce = await OAuthNonce.findAndDelete(read, persis, rocketChatUserId);
            if (storedNonce === null || storedNonce !== parsed.nonce) {
                this.app.getLogger().warn('Authentication rejected — nonce mismatch.');
                return null;
            }

            return { rocketChatUserId, type, accessCode: request.query.code as string };
        } catch (error) {
            this.app.getLogger().error(
                `Authentication — state validation failed: ${(error as Error)?.message ?? String(error)}`
            );
            return null;
        }
    }

    private async fetchAppEnvironment(read: IRead): Promise<{ aadTenantId: string; aadClientId: string; aadClientSecret: string; authEndpointUrl: string } | null> {
        try {
            const [aadTenantId, aadClientId, aadClientSecret, authEndpointUrl] = await Promise.all([
                read.getEnvironmentReader().getSettings().getValueById(AppSetting.AadTenantId),
                read.getEnvironmentReader().getSettings().getValueById(AppSetting.AadClientId),
                read.getEnvironmentReader().getSettings().getValueById(AppSetting.AadClientSecret),
                getRocketChatAppEndpointUrl(this.app.getAccessors(), AuthenticationEndpointPath)
            ]);

            return { aadTenantId, aadClientId, aadClientSecret, authEndpointUrl };
        } catch (error) {
            this.app.getLogger().error(
                `Authentication — settings fetch failed: ${(error as Error)?.message ?? String(error)}`
            );
            return null;
        }
    }

    private async exchangeAndValidateTokens(
        http: IHttp, read: IRead, persis: IPersistence, accessCode: string, authEndpointUrl: string,
        aadTenantId: string, aadClientId: string, aadClientSecret: string, type: 'bot' | 'normal', rocketChatUserId: string
    ) {
        try {
            const response = await getUserAccessTokenAsync(
                http, accessCode, authEndpointUrl, aadTenantId, aadClientId, aadClientSecret, type,
            );

            if (!response.refreshToken) {
                throw new Error('Token exchange did not return a refresh token.');
            }

            const { accessToken: userAccessToken, refreshToken, expiresIn, extExpiresIn } = response;

            const teamsUserProfile = await getUserProfileAsync(http, userAccessToken);
            const teamsUserId = teamsUserProfile.id;

            const appUser = await read.getUserReader().getAppUser(this.app.getID());

            const existingMapping = await UserMapping.findByTeamsUserId(read, teamsUserId);
            if (existingMapping !== null && existingMapping.rocketChatUserId !== rocketChatUserId) {
                if (appUser && existingMapping.rocketChatUserId === appUser.id) {
                    return { conflictError: "This Teams account is used by the app user and cannot be linked to a personal Rocket.Chat account." };
                }

                const existingUser = await read.getUserReader().getById(existingMapping.rocketChatUserId);
                if (!existingUser || !existingUser.isEnabled) {
                    // If the user linked is disabled or deleted, clean up the mapping
                    await Promise.all([
                        UserRegistration.delete(persis, existingMapping.rocketChatUserId),
                        UserMapping.delete(read, persis, existingMapping.rocketChatUserId),
                        LoginMessage.delete(persis, existingMapping.rocketChatUserId),
                    ]);
                    this.app.getLogger().log(
                        `[Teams Bridge] Cleaned up orphaned Teams link mapping for deleted/inactive user ${existingMapping.rocketChatUserId} during new login.`
                    );
                } else {
                    return { conflictError: `This Teams account is already linked to another Rocket.Chat user (@${existingUser.username}). Please log out from Teams on the other account first.` };
                }
            }

            return { userAccessToken, refreshToken, expiresIn, extExpiresIn, teamsUserId, appUser };
        } catch (error) {
            this.app.getLogger().error(
                `Authentication — token exchange/profile failed: ${(error as Error)?.message ?? String(error)}`
            );
            return null;
        }
    }

    private async persistAndSubscribe(
        read: IRead, persis: IPersistence, http: IHttp, rocketChatUserId: string, type: 'bot' | 'normal', tokenData: any
    ) {
        try {
            const { userAccessToken, refreshToken, expiresIn, extExpiresIn, teamsUserId, appUser } = tokenData;

            await Promise.all([
                UserRegistration.persist(
                    persis, rocketChatUserId, userAccessToken, refreshToken, expiresIn, extExpiresIn
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

            if (type === 'bot') {
                await subscribeToAllMessagesForOneUserAsync({
                    http, read, persis, rocketChatUserId,
                    subscriberEndpointUrl, teamsUserId, userAccessToken,
                    renewIfExists: true, forceRenew: true,
                });
            }

            return true;
        } catch (error) {
            this.app.getLogger().error(
                `Authentication — persistence/subscription failed: ${(error as Error)?.message ?? String(error)}`
            );
            return false;
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
