import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiSubscriptionOperationUrl } from '../Const';

export const deleteSubscriptionAsync = async (http: IHttp, subscriptionId: string, userAccessToken: string): Promise<void> => {
	const url = getGraphApiSubscriptionOperationUrl(subscriptionId);

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.del(url, httpRequest);

	if (response.statusCode !== HttpStatusCode.NO_CONTENT) {
		throw new Error(`Delete subscription failed with http status code ${response.statusCode}.`);
	}
};
