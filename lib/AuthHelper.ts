import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import type { TeamsBridgeApp } from "../TeamsBridgeApp";
import { AppSetting } from "../config/Settings";
import { renewUserAccessTokenAsync } from "./MicrosoftGraphApi";
import {
    persistUserAccessTokenAsync,
    retrieveAllUserRegistrationsAsync,
    retrieveUserAccessTokenDataAsync,
    saveLoginMessageSentStatus,
    UserRegistrationModel,
} from "./PersistHelper";

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
                    aadClientSecret
                );
                await persistUserAccessTokenAsync(
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
        await saveLoginMessageSentStatus({
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

    const registration = await retrieveUserAccessTokenDataAsync({
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

export const getAllUsersAccessTokensAsync = async (options: {
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
    forceRefresh?: boolean;
}) => {
    const { read } = options;

    const registrations = await retrieveAllUserRegistrationsAsync(read);

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
