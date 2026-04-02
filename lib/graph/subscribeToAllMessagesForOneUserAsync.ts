import type { IHttp, IHttpRequest, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';
import * as ms from 'ms';

import { getGraphApiSubscriptionUrl, SupportedNotificationChangeTypes, SubscriptionMaxExpireTimeInSecond, RegistrationAutoRenewInterval } from '../Const';
import { WebhookSecret } from '../PersistHelper';
import { getNotificationEndpointUrl } from '../UrlHelper';
import { deleteSubscriptionAsync } from './deleteSubscriptionAsync';
import { listSubscriptionsAsync } from './listSubscriptionsAsync';
import { renewSubscriptionAsync } from './renewSubscriptionAsync';
import type { SubscriptionResponse } from './types';

export const subscribeToAllMessagesForOneUserAsync = async (options: {
	http: IHttp;
	read: IRead;
	persis: IPersistence;
	rocketChatUserId: string;
	teamsUserId: string;
	subscriberEndpointUrl: string;
	userAccessToken: string;
	expirationDateTime?: Date;
	renewIfExists?: boolean;
	forceRenew?: boolean;
}): Promise<SubscriptionResponse | undefined> => {
	const {
		rocketChatUserId,
		teamsUserId,
		subscriberEndpointUrl,
		userAccessToken,
		expirationDateTime: inputExpirationDateTime,
		renewIfExists = true,
		forceRenew = false,
		http,
		read,
		persis,
	} = options;

	let expirationDateTime = inputExpirationDateTime;
	if (!expirationDateTime) {
		expirationDateTime = new Date();
		expirationDateTime.setSeconds(expirationDateTime.getSeconds() + SubscriptionMaxExpireTimeInSecond);
	}

	const url = getGraphApiSubscriptionUrl();

	const notificationUrl = getNotificationEndpointUrl({
		subscriberEndpoint: subscriberEndpointUrl,
		rocketChatUserId,
	});

	const clientState = await WebhookSecret.getSubscriptionStateHash(read.getPersistenceReader(), persis, { rocketChatUserId });
	const body = {
		changeType: SupportedNotificationChangeTypes.join(','),
		notificationUrl,
		resource: `/users/${teamsUserId}/chats/getAllMessages`,
		includeResourceData: false,
		expirationDateTime: expirationDateTime.toISOString(),
		clientState,
	};

	if (renewIfExists) {
		const existingSubscriptions = (await listSubscriptionsAsync(http, userAccessToken, notificationUrl)) || [];

		if (existingSubscriptions.length > 0) {
			if (existingSubscriptions.length > 1) {
				await Promise.all(existingSubscriptions.slice(1).map((sub) => deleteSubscriptionAsync(http, sub.id, userAccessToken)));
			}
			const hasClientState = new URL(existingSubscriptions[0].notificationUrl).searchParams.has('hasClientState');
			if (!hasClientState) {
				await deleteSubscriptionAsync(http, existingSubscriptions[0].id, userAccessToken);
			} else {
				const existingSub = existingSubscriptions[0];
				const currentExpireTime = new Date(existingSub.expirationDateTime);
				const now = new Date();
				const nextUpdateTime = new Date(now.getTime() + ms(RegistrationAutoRenewInterval));

				const timeLeftAtNextUpdate = currentExpireTime.getTime() - nextUpdateTime.getTime();
				const threshold = ms(RegistrationAutoRenewInterval) / 2;
				const shouldRenew = timeLeftAtNextUpdate <= threshold;

				if (!shouldRenew && !forceRenew) {
					return;
				}
				return await renewSubscriptionAsync(http, existingSubscriptions[0].id, userAccessToken, expirationDateTime, clientState);
			}
		}
	}

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
		content: JSON.stringify(body),
	};

	const response = await http.post(url, httpRequest);

	if (response.statusCode === HttpStatusCode.CREATED) {
		const responseBody = response.data;
		if (responseBody === undefined) {
			throw new Error('Subscribe to notification for user failed!');
		}

		const result: SubscriptionResponse = {
			subscriptionId: responseBody.id,
			expirationTime: new Date(responseBody.expirationDateTime),
		};

		return result;
	}
	throw new Error(
		`Subscribe to notification for user failed with http status code ${response.statusCode}.\nReceived: ${JSON.stringify(response.data, null, 2)}`,
	);
};
