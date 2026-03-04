import { IHttp } from "@rocket.chat/apps-engine/definition/accessors";
import { listSubscriptionsAsync } from './listSubscriptionsAsync';
import { deleteSubscriptionAsync } from './deleteSubscriptionAsync';

export const deleteAllSubscriptions = async (http: IHttp, userAccessToken: string, notificationUrl: string) => {
    const subscriptionsId = (
        await listSubscriptionsAsync(
            http,
            userAccessToken,
            notificationUrl,
        )
    )?.map((subscription) => (subscription as any).id);
    if (subscriptionsId) {
        for (const subscriptionId of subscriptionsId) {
            try {
                await deleteSubscriptionAsync(
                    http,
                    subscriptionId,
                    userAccessToken
                );
            } catch (error) {
                console.error(
                    `Error during delete subscription, will ignore and continue. ${error}`
                );
            }
        }
    }
};
