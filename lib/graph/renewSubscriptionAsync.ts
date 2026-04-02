import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiSubscriptionOperationUrl, SubscriptionMaxExpireTimeInSecond } from '../Const';
import type { SubscriptionResponse } from './types';

export const renewSubscriptionAsync = async (
	http: IHttp,
	subscriptionId: string,
	userAccessToken: string,
	expirationDateTime?: Date,
	clientState?: string,
): Promise<SubscriptionResponse | undefined> => {
	if (!expirationDateTime) {
		expirationDateTime = new Date();
		expirationDateTime.setSeconds(expirationDateTime.getSeconds() + SubscriptionMaxExpireTimeInSecond);
	}

	const url = getGraphApiSubscriptionOperationUrl(subscriptionId);

	const body = {
		expirationDateTime: expirationDateTime.toISOString(),
		...(clientState && { clientState }),
	};

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
		content: JSON.stringify(body),
	};

	const response = await http.patch(url, httpRequest);

	if (response.statusCode === HttpStatusCode.OK) {
		const responseBody = response.data;
		if (responseBody === undefined) {
			throw new Error('Renew subscription failed!');
		}

		const result: SubscriptionResponse = {
			subscriptionId: responseBody.id,
			expirationTime: new Date(responseBody.expirationDateTime),
		};

		return result;
	}
	console.error(`Renew subscription failed with http status code ${response.statusCode}.\nReceived: ${JSON.stringify(response.data, null, 2)}`);
};
