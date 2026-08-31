import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiChannelMembersUrl } from '../Const';
import type { TeamsChatMember } from './getTeamsChatMembersAsync';

// Lists the members of a Team channel, in the same shape as chat members so the
// mapping UIs can consume either source.
export const getChannelMembersAsync = async (
	http: IHttp,
	teamId: string,
	channelId: string,
	userAccessToken: string,
): Promise<{ members: TeamsChatMember[]; nextLink?: string } | null> => {
	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(getGraphApiChannelMembersUrl(teamId, channelId), httpRequest);
	if (response.statusCode !== HttpStatusCode.OK) {
		return null;
	}
	const raw = (response.data?.value ?? []) as any[];
	const members = raw
		.map((m) => ({
			userId: m.userId ?? '',
			displayName: m.displayName ?? m.email ?? m.userId ?? 'Unknown',
		}))
		.filter((m) => m.userId);
	return { members };
};
