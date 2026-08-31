import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiTeamChannelsUrl } from '../Const';

export interface TeamChannel {
	id: string; // 19:...@thread.tacv2
	displayName: string;
}

// Lists the channels of one Team.
export const listTeamChannelsAsync = async (http: IHttp, teamId: string, userAccessToken: string): Promise<TeamChannel[] | null> => {
	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(getGraphApiTeamChannelsUrl(teamId), httpRequest);
	if (response.statusCode !== HttpStatusCode.OK) {
		return null;
	}
	const raw = (response.data?.value ?? []) as any[];
	return raw.map((c) => ({ id: c.id ?? '', displayName: c.displayName ?? c.id ?? 'Unknown channel' })).filter((c) => c.id);
};
