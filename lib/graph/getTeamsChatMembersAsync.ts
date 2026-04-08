import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiChatMemberUrl } from '../Const';

export interface TeamsChatMember {
	userId: string;
	displayName: string;
}

export interface GetTeamsChatMembersResult {
	members: TeamsChatMember[];
	nextLink?: string;
}

export const getTeamsChatMembersAsync = async (
	http: IHttp,
	threadId: string,
	userAccessToken: string,
	options: { pageUrl?: string } = {},
): Promise<GetTeamsChatMembersResult | null> => {
	const url = options.pageUrl ?? getGraphApiChatMemberUrl(threadId);
	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(url, httpRequest);

	if (response.statusCode !== HttpStatusCode.OK) {
		return null;
	}

	const userList = (response.data?.value ?? []) as any[];
	const members = userList.map((u) => ({
		userId: u.userId ?? '',
		displayName: u.displayName ?? u.userId ?? 'Unknown',
	}));

	const nextLink: string | undefined = response.data?.['@odata.nextLink'] ?? undefined;

	return { members, nextLink };
};
