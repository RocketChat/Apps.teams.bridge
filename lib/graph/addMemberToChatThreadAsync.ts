import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiChatMemberUrl } from '../Const';

export type AddMemberResult = { status: 'added' } | { status: 'already_member' } | { status: 'failed'; statusCode: number };

export const addMemberToChatThreadAsync = async (http: IHttp, threadId: string, teamsUserId: string, userAccessToken: string): Promise<AddMemberResult> => {
	const headers = {
		'Content-Type': 'application/json',
		Authorization: `Bearer ${userAccessToken}`,
	};

	const filterUrl = `${getGraphApiChatMemberUrl(threadId)}?$filter=microsoft.graph.aadUserConversationMember/userId eq '${teamsUserId}'`;
	const checkResponse = await http.get(filterUrl, { headers });
	if (checkResponse.statusCode === HttpStatusCode.OK) {
		const members = (checkResponse.data?.value ?? []) as any[];
		if (members.length > 0) {
			return { status: 'already_member' };
		}
	}

	const url = getGraphApiChatMemberUrl(threadId);

	const body = {
		'@odata.type': '#microsoft.graph.aadUserConversationMember',
		roles: ['owner'],
		'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${teamsUserId}`,
		visibleHistoryStartDateTime: '0001-01-01T00:00:00Z',
	};

	const httpRequest: IHttpRequest = {
		headers,
		content: JSON.stringify(body),
	};

	const response = await http.post(url, httpRequest);

	if (response.statusCode === HttpStatusCode.CREATED) {
		return { status: 'added' };
	}

	// 409 Conflict = user is already a member of the chat
	if (response.statusCode === 409) {
		return { status: 'already_member' };
	}

	return { status: 'failed', statusCode: response.statusCode };
};
