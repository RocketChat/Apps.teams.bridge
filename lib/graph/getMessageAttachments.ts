import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiMessageUrl } from '../Const';

export const getMessageAttachments = async ({
	http,
	messageId,
	threadId,
	userAccessToken,
}: {
	http: IHttp;
	userAccessToken: string;
	messageId: string;
	threadId: string;
}) => {
	const url = getGraphApiMessageUrl(threadId, messageId, false);

	const httpRequest: IHttpRequest = {
		headers: {
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(url, httpRequest);

	if (response.statusCode === HttpStatusCode.OK) {
		const { attachments = [] } = response.data || {};
		return attachments as any[];
	}
	console.error(`Get Teams message by ID failed with http status code ${response.statusCode}.`);
	return [];
};
