import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiMessageDeleteUrl } from '../Const';

export const deleteTextMessageInChatThreadAsync = async (
	http: IHttp,
	teamsUserId: string,
	messageId: string,
	threadId: string,
	userAccessToken: string,
): Promise<void> => {
	const url = getGraphApiMessageDeleteUrl(teamsUserId, threadId, messageId);

	const httpRequest: IHttpRequest = {
		headers: {
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.post(url, httpRequest);

	if (response.statusCode === HttpStatusCode.NO_CONTENT) {
	} else {
		throw new Error(`Delete message in chat thread failed with http status code ${response.statusCode}.`);
	}
};
