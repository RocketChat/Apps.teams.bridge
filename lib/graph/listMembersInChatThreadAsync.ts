import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiChatMemberUrl } from '../Const';

export const listMembersInChatThreadAsync = async (http: IHttp, threadId: string, userAccessToken: string): Promise<{ id: string; userId: string }[]> => {
	const url = getGraphApiChatMemberUrl(threadId);
	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(url, httpRequest);

	if (response.statusCode === HttpStatusCode.OK) {
		const responseBody = response.data;
		if (responseBody === undefined) {
			throw new Error('List members in chat thread failed!');
		}

		const userList = responseBody.value as any[];
		const result: { id: string; userId: string }[] = [];

		for (const user of userList) {
			result.push({
				id: user.id,
				userId: user.userId,
			});
		}

		return result;
	}
	throw new Error(`List members in chat thread failed with http status code ${response.statusCode}.`);
};
