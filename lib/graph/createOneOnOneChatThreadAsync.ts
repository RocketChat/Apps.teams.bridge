import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiChatUrl, GraphApiBaseUrl } from '../Const';
import type { CreateThreadResponse } from './types';

export const createOneOnOneChatThreadAsync = async (
	http: IHttp,
	senderUserTeamsId: string,
	receiverUserTeamsId: string,
	userAccessToken: string,
): Promise<CreateThreadResponse> => {
	const url = getGraphApiChatUrl();

	const body = {
		chatType: 'oneOnOne',
		members: [senderUserTeamsId, receiverUserTeamsId].map((userId) => ({
			'@odata.type': '#microsoft.graph.aadUserConversationMember',
			roles: ['owner'],
			'user@odata.bind': `${GraphApiBaseUrl}/v1.0/users('${userId}')`,
		})),
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
			throw new Error('Create one on one chat thread failed!');
		}

		const result: CreateThreadResponse = {
			threadId: responseBody.id,
		};

		return result;
	}
	throw new Error(`Create one on one chat thread failed with http status code ${response.statusCode}.\n Received: ${JSON.stringify(response.data, null, 2)}`);
};
