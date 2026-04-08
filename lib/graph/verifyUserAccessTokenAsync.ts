import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiProfileUrl } from '../Const';

export const verifyUserAccessTokenAsync = async (http: IHttp, userAccessToken: string): Promise<boolean> => {
	const url = getGraphApiProfileUrl();
	const httpRequest: IHttpRequest = {
		headers: {
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	try {
		const response = await http.get(url, httpRequest);
		return response.statusCode === HttpStatusCode.OK;
	} catch (e) {
		return false;
	}
};
