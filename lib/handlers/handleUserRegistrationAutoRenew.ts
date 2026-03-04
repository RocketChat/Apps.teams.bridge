import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { subscribeToAllMessagesForOneUserAsync } from "../MicrosoftGraphApi";
import { UserMapping, UserRegistration } from "../PersistHelper";

export const handleUserRegistrationAutoRenewAsync = async (options: {
    subscriberEndpointUrl: string;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp,
}): Promise<void> => {
    const { http, persistence, read, subscriberEndpointUrl, app } = options;

    const allRegistrations = await UserRegistration.findAll(read);

    if (allRegistrations) {
        const errorUserIds: string[] = [];
        for (const registration of allRegistrations) {
            try {
                const userAccessToken = await getUserAccessTokenAsync({
                    app,
                    http,
                    persistence,
                    read,
                    rocketChatUserId: registration.rocketChatUserId,
                });

                if (!userAccessToken) {
                    errorUserIds.push(registration.rocketChatUserId);
                    continue;
                }

                const user = await UserMapping.findByRCUserId(
                    read,
                    registration.rocketChatUserId
                );

                if (!user) {
                    throw new Error(
                        `User record for user ${registration.rocketChatUserId} not found!`
                    );
                }

                await subscribeToAllMessagesForOneUserAsync({
                    read,
                    http,
                    persis: persistence,
                    rocketChatUserId: user.rocketChatUserId,
                    subscriberEndpointUrl,
                    teamsUserId: user.teamsUserId,
                    userAccessToken,
                    renewIfExists: true,
                    forceRenew: false,
                });
            } catch (error) {
                console.error(
                    `Error during renew registration for user ${registration.rocketChatUserId}. Ignore this error and continue. Error: ${error}`
                );
            }
        }
        if (errorUserIds.length) {
            app.getLogger().error(`Could not refresh user access token for users: ${errorUserIds.join(', ')}`)
        }
    }
};
