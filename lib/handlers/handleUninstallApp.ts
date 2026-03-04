import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getAllUsersAccessTokensAsync } from "../AuthHelper";
import { deleteAllSubscriptions } from "../MicrosoftGraphApi";
import { getNotificationEndpointUrl } from "../UrlHelper";

export const handleUninstallApp = async (options: {
    read: IRead;
    http: IHttp;
    modify: IModify;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}) => {
    const { modify, app } = options;
    try {
        await deleteAllUsersSubscriptions(options),
        await app.deleteAppUsers(modify);
    } catch (error) {
        console.error(`Error during app uninstallation: ${error.message}`);
    }
};

const deleteAllUsersSubscriptions = async (options: {
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
}) => {
    const { read, persistence, http, app } = options;
    const allRegisteredUsers = await getAllUsersAccessTokensAsync({
        read,
        http,
        app,
        persistence,
    });

    if (!allRegisteredUsers) {
        return;
    }

    const batchSize = 10;
    for (let i = 0; i < allRegisteredUsers.length; i += batchSize) {
        const batch = allRegisteredUsers.slice(i, i + batchSize);
        const deletePromises = batch.map(async ({ accessToken, rocketChatUserId }) => {
            try {
                const notificationUrl = await getNotificationEndpointUrl({
                    appAccessors: app.getAccessors(),
                    rocketChatUserId: rocketChatUserId,
                });
                if (accessToken) {
                    await deleteAllSubscriptions(
                        http,
                        accessToken,
                        notificationUrl
                    );
                }
            } catch (error) {
                console.error(
                    `Error deleting subscriptions for user: ${error.message}`
                );
            }
        });
        await Promise.all(deletePromises); // Wait for the current batch to complete
    }
};
