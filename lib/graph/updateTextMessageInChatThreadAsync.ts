import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiMessageUrl } from '../Const';

export const updateTextMessageInChatThreadAsync = async ({
	http,
	textMessage,
	messageType,
	messageId,
	threadId,
	userAccessToken,
	attachments,
}: {
	http: IHttp;
	textMessage: string;
	messageType: 'text' | 'html';
	messageId: string;
	threadId: string;
	userAccessToken: string;
	attachments?: any[];
}): Promise<void> => {
	const url = getGraphApiMessageUrl(threadId, messageId, true);

	const body: any = {
		body: {
			content: textMessage,
			contentType: messageType,
		},
	};

	if (attachments && attachments.length > 0) {
		body.attachments = attachments;
	}

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
		content: JSON.stringify(body),
	};

	const response = await http.patch(url, httpRequest);

	if (response.statusCode === HttpStatusCode.NO_CONTENT) {
	} else {
		throw new Error(`Update message in chat thread failed with http status code ${response.statusCode}.`);
	}
};
