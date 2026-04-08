import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiChatMemberRemoveUrl } from '../Const';

export const removeMemberFromChatThreadAsync = async (http: IHttp, threadId: string, membershipId: string, userAccessToken: string): Promise<void> => {
	const url = getGraphApiChatMemberRemoveUrl(threadId, membershipId);
	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.del(url, httpRequest);

	if (response.statusCode !== HttpStatusCode.NO_CONTENT) {
		throw new Error(`Remove member from chat thread failed with http status code ${response.statusCode}.`);
	}
};
