import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiMessageUrl } from '../Const';
import type { SendMessageResponse } from './types';

export const sendFileMessageToChatThreadAsync = async (
	http: IHttp,
	textMessage: string,
	fileName: string,
	shareLink: string,
	threadId: string,
	userAccessToken: string,
	teamId?: string, // set for channel-linked rooms
): Promise<SendMessageResponse> => {
	const url = getGraphApiMessageUrl(threadId, undefined, false, teamId);

	const body = {
		body: {
			content: `${textMessage} <a href=\"${shareLink}\" title=\"${shareLink}\" target=\"_blank\" rel=\"noreferrer noopener\">${fileName}</a>`,
			contentType: 'html',
		},
	};

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
			throw new Error('Send file message to chat thread failed!');
		}

		const result: SendMessageResponse = {
			messageId: responseBody.id,
		};

		return result;
	}
	throw new Error(`Send file message to chat thread failed with http status code ${response.statusCode}.`);
};
