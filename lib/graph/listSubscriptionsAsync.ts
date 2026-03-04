import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { getGraphApiSubscriptionUrl } from "../Const";
import { SubscriptionsResponse } from './types';

export const listSubscriptionsAsync = async (
    http: IHttp,
    userAccessToken: string,
    notificationUrl: string,
): Promise<SubscriptionsResponse['value'] | undefined> => {
    const url = getGraphApiSubscriptionUrl();

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error("List subscriptions failed!");
        }

        const subscriptions = response.data?.value as SubscriptionsResponse['value'];
        const urlObj = new URL(notificationUrl);
        const pathWithQuery = urlObj.pathname + urlObj.search;
        return subscriptions.filter((subscription) => subscription.notificationUrl.includes(pathWithQuery));
    } else {
        console.error(
            `List subscriptions failed with http status code ${response.statusCode}. \nReceived: ${JSON.stringify(response.data, null, 2)}`
        );
        return;
    }
};
