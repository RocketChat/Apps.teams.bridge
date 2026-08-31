import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiJoinedTeamsUrl } from '../Const';

export interface JoinedTeam {
	id: string;
	displayName: string;
}

// Lists the Teams (the org kind) the signed-in bot account belongs to.
export const listJoinedTeamsAsync = async (http: IHttp, userAccessToken: string): Promise<JoinedTeam[] | null> => {
	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(getGraphApiJoinedTeamsUrl(), httpRequest);
	if (response.statusCode !== HttpStatusCode.OK) {
		return null;
	}
	const raw = (response.data?.value ?? []) as any[];
	return raw.map((t) => ({ id: t.id ?? '', displayName: t.displayName ?? t.id ?? 'Unknown team' })).filter((t) => t.id);
};
