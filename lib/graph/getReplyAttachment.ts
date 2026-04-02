import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiMessageUrl } from '../Const';

export const getReplyAttachment = async ({
	http,
	parentMessageId,
	threadId,
	userAccessToken,
}: {
	http: IHttp;
	userAccessToken: string;
	parentMessageId: string;
	threadId: string;
}) => {
	const url = getGraphApiMessageUrl(threadId, parentMessageId, false);

	const httpRequest: IHttpRequest = {
		headers: {
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(url, httpRequest);

	if (response.statusCode === HttpStatusCode.OK) {
		const responseBody = response.data || {};
		const { id, from, body } = responseBody;

		if (!id || !from?.user || !body?.content) {
			return;
		}
		return {
			id,
			contentType: 'messageReference',
			content: JSON.stringify({
				messageId: id,
				messagePreview: body.content,
				messageSender: {
					user: from.user,
				},
			}),
		};
	}
	console.error(`Get Teams message by ID failed with http status code ${response.statusCode}.`);
};
