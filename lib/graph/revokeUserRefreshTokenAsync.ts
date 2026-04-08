import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiRevokeRefreshTokenUrl } from '../Const';

export const revokeUserRefreshTokenAsync = async (http: IHttp, userAccessToken: string): Promise<void> => {
	const url = getGraphApiRevokeRefreshTokenUrl();
	const httpRequest: IHttpRequest = {
		headers: {
			Authorization: `Bearer ${userAccessToken}`,
			'Content-Type': 'application/json',
			Accept: 'application/json',
		},
		content: '{}',
	};

	const response = await http.post(url, httpRequest);

	if (response.statusCode !== HttpStatusCode.OK) {
		throw new Error(`Revoke user refresh token failed with http status code ${response.statusCode}.\nReceived: ${JSON.stringify(response.data, null, 2)}`);
	}
};
