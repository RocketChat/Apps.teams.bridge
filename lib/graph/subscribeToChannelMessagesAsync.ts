import type { IHttp, IHttpRequest, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiSubscriptionUrl, SupportedNotificationChangeTypes, SubscriptionMaxExpireTimeInSecond } from '../Const';
import { WebhookSecret } from '../PersistHelper';
import { deleteSubscriptionAsync } from './deleteSubscriptionAsync';
import { listSubscriptionsAsync } from './listSubscriptionsAsync';
import { renewSubscriptionAsync } from './renewSubscriptionAsync';
import type { SubscriptionResponse } from './types';

// Subscribes to new messages in one Team channel. Mirrors subscribeToAllMessagesForOneUserAsync,
// but per-channel: resource is /teams/{teamId}/channels/{channelId}/messages and the webhook URL
// carries a channelId param placed BEFORE hasClientState so the chat subscription's URL filter
// (a substring match) can never collide with channel subscriptions, and vice versa.
export const subscribeToChannelMessagesAsync = async (options: {
	http: IHttp;
	read: IRead;
	persis: IPersistence;
	rocketChatUserId: string; // the app user's RC id (notification receiver)
	teamId: string;
	channelId: string;
	subscriberEndpointUrl: string;
	userAccessToken: string;
	renewIfExists?: boolean;
}): Promise<SubscriptionResponse | undefined> => {
	const { http, read, persis, rocketChatUserId, teamId, channelId, subscriberEndpointUrl, userAccessToken, renewIfExists = true } = options;

	const expirationDateTime = new Date();
	expirationDateTime.setSeconds(expirationDateTime.getSeconds() + SubscriptionMaxExpireTimeInSecond);

	const notificationUrl = `${subscriberEndpointUrl}?userId=${rocketChatUserId}&channelId=${encodeURIComponent(channelId)}&hasClientState=1`;
	const clientState = await WebhookSecret.getSubscriptionStateHash(read.getPersistenceReader(), persis, { rocketChatUserId });

	if (renewIfExists) {
		const existing = (await listSubscriptionsAsync(http, userAccessToken, notificationUrl)) || [];
		if (existing.length > 0) {
			if (existing.length > 1) {
				await Promise.all(existing.slice(1).map((sub) => deleteSubscriptionAsync(http, sub.id, userAccessToken)));
			}
			return renewSubscriptionAsync(http, existing[0].id, userAccessToken, expirationDateTime, clientState);
		}
	}

	const body = {
		changeType: SupportedNotificationChangeTypes.join(','),
		notificationUrl,
		resource: `/teams/${teamId}/channels/${channelId}/messages`,
		includeResourceData: false,
		expirationDateTime: expirationDateTime.toISOString(),
		clientState,
	};

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
		content: JSON.stringify(body),
	};

	const response = await http.post(getGraphApiSubscriptionUrl(), httpRequest);
	if (response.statusCode === HttpStatusCode.CREATED) {
		const responseBody = response.data;
		if (responseBody === undefined) {
			throw new Error('Subscribe to channel messages failed!');
		}
		return { subscriptionId: responseBody.id, expirationTime: new Date(responseBody.expirationDateTime) };
	}
	throw new Error(
		`Subscribe to channel messages failed with http status code ${response.statusCode}.\nReceived: ${JSON.stringify(response.data, null, 2)}`,
	);
};
