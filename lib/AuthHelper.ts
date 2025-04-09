import { IHttp, IPersistence, IPersistenceRead, IRead } from "@rocket.chat/apps-engine/definition/accessors";
import type { TeamsBridgeApp } from "../TeamsBridgeApp";
import { RocketChatAssociationModel, RocketChatAssociationRecord } from "@rocket.chat/apps-engine/definition/metadata";
import { AppSetting } from "../config/Settings";
import { renewUserAccessTokenAsync } from "./MicrosoftGraphApi";
import { MiscKeys, persistUserAccessTokenAsync, retrieveUserAccessTokenDataAsync, saveLoginMessageSentStatus, UserRegistrationModel } from "./PersistHelper";

export const getUserAccessTokenAsync = async (options: {
    read: IRead;
    persistence: IPersistence;
    rocketChatUserId: string;
    http: IHttp;
    app: TeamsBridgeApp;
    forceRefresh?: boolean;
}): Promise<string | null> => {
    const { read, rocketChatUserId, http, persistence, app, forceRefresh = false } = options;

    const data = await retrieveUserAccessTokenDataAsync({
        rocketChatUserId,
        read,
    });

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    if (!data || !data.expires || epochInSecond > data.expires || forceRefresh) {
        if (data?.refreshToken) {
            const [aadClientId, aadTenantId, aadClientSecret] = await Promise.all([
                app.getSettingValueById(AppSetting.AadClientId),
                app.getSettingValueById(AppSetting.AadTenantId),
                app.getSettingValueById(AppSetting.AadClientSecret),
            ])

            try {
                const response = await renewUserAccessTokenAsync(http, data.refreshToken, aadTenantId, aadClientId, aadClientSecret);
                await persistUserAccessTokenAsync(
                    persistence,
                    rocketChatUserId,
                    response.accessToken,
                    response.refreshToken as string,
                    response.expiresIn,
                    response.extExpiresIn
                );
                return response.accessToken;
            } catch (error) {
                app.getLogger().error(`Failed to renew access token for user ${rocketChatUserId}`, error);
            }
        }
        await saveLoginMessageSentStatus({
            persistence,
            rocketChatUserId,
            wasSent: false,
        });
        return null;
    }

    return data.accessToken;
};
