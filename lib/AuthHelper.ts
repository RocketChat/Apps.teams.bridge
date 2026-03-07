import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import type { TeamsBridgeApp } from "../TeamsBridgeApp";
import { AppSetting } from "../config/Settings";
import { renewUserAccessTokenAsync } from "./MicrosoftGraphApi";
import { AppToken, LoginMessage, UserRegistration } from "./PersistHelper";
import type { UserRegistrationModel } from "./PersistHelper";

export const getAccessTokenForRegistration = async (options: {
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
    forceRefresh?: boolean;
    registration: UserRegistrationModel;
}): Promise<string | null> => {
    const {
        registration,
        http,
        persistence,
        app,
        forceRefresh = false,
    } = options;

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    const appUser = await app.getAccessors().reader.getUserReader().getAppUser();
    const userType = appUser?.id === registration.rocketChatUserId ? 'bot' : 'normal';

    if (
        !registration ||
        !registration.expires ||
        epochInSecond > registration.expires ||
        forceRefresh
    ) {
        if (registration?.refreshToken) {
            const [aadClientId, aadTenantId, aadClientSecret] =
                await Promise.all([
                    app.getSettingValueById(AppSetting.AadClientId),
                    app.getSettingValueById(AppSetting.AadTenantId),
                    app.getSettingValueById(AppSetting.AadClientSecret),
                ]);

            try {
                const response = await renewUserAccessTokenAsync(
                    http,
                    registration.refreshToken,
                    aadTenantId,
                    aadClientId,
                    aadClientSecret,
                    userType,
                );
                await UserRegistration.persist(
                    persistence,
                    registration.rocketChatUserId,
                    response.accessToken,
                    response.refreshToken as string,
                    response.expiresIn,
                    response.extExpiresIn
                );
                return response.accessToken;
            } catch (error) {
                app.getLogger().error(
                    `Failed to renew access token for user ${registration.rocketChatUserId}`,
                    error
                );
            }
        }
        await LoginMessage.save({
            persistence,
            rocketChatUserId: registration.rocketChatUserId,
            wasSent: false,
        });
        return null;
    }

    return registration.accessToken;
};

export const getUserAccessTokenAsync = async (options: {
    read: IRead;
    persistence: IPersistence;
    rocketChatUserId: string;
    http: IHttp;
    app: TeamsBridgeApp;
    forceRefresh?: boolean;
}): Promise<string | null> => {
    const { read, rocketChatUserId } = options;

    const registration = await UserRegistration.findByRCUserId({
        rocketChatUserId,
        read,
    });

    if (!registration) {
        return null;
    }

    return getAccessTokenForRegistration({
        registration,
        ...options,
    });
};

export const getAppAccessTokenAsync = async (options: {
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
}): Promise<string | null> => {
    const { read, persistence, http, app } = options;

    const EXPIRY_BUFFER_SECONDS = 60;
    const epochNow = Math.round(Date.now() / 1000);

    // Check persistence first — avoid an HTTP round-trip if the token is still valid.
    const cached = await AppToken.find(read);
    if (cached && cached.expires && epochNow < cached.expires - EXPIRY_BUFFER_SECONDS) {
        return cached.accessToken;
    }

    const [aadClientId, aadTenantId, aadClientSecret] = await Promise.all([
        app.getSettingValueById(AppSetting.AadClientId),
        app.getSettingValueById(AppSetting.AadTenantId),
        app.getSettingValueById(AppSetting.AadClientSecret),
    ]);

    try {
        const response = await http.post(
            `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/token`,
            {
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                content: `client_id=${encodeURIComponent(aadClientId)}&client_secret=${encodeURIComponent(aadClientSecret)}&grant_type=client_credentials&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default`,
            }
        );

        if (!response?.data?.access_token) {
            app.getLogger().error('getAppAccessTokenAsync: no access_token in response', response?.data);
            return null;
        }

        const accessToken = response.data.access_token as string;
        const expiresIn = (response.data.expires_in as number) ?? 3600;
        await AppToken.persist(persistence, accessToken, epochNow + expiresIn);
        return accessToken;
    } catch (error) {
        app.getLogger().error('getAppAccessTokenAsync: failed to fetch app token', error);
        return null;
    }
};

export const getAllUsersAccessTokensAsync = async (options: {
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
    forceRefresh?: boolean;
}) => {
    const { read } = options;

    const registrations = await UserRegistration.findAll(read);

    if (!registrations) {
        return null;
    }

    const batchSize = 10;
    const results: { accessToken: string | null; rocketChatUserId: string }[] = [];

    for (let i = 0; i < registrations.length; i += batchSize) {
        const batch = registrations.slice(i, i + batchSize);

        const batchResults = await Promise.all(
            batch.map((registration) =>
                Promise.all([
                    getAccessTokenForRegistration({
                        registration,
                        ...options,
                    }),
                    Promise.resolve(registration.rocketChatUserId),
                ])
            )
        );

        results.push(
            ...batchResults.map(([accessToken, rocketChatUserId]) => ({
                accessToken,
                rocketChatUserId,
            }))
        );
    }

    return results;
};
