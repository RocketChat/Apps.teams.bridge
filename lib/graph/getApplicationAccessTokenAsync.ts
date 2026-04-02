import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getMicrosoftTokenUrl } from '../Const';
import type { TokenResponse } from './types';

export const getApplicationAccessTokenAsync = async (
	http: IHttp,
	aadTenantId: string,
	aadClientId: string,
	aadClientSecret: string,
): Promise<TokenResponse> => {
	const requestBody =
		'scope=https://graph.microsoft.com/.default&grant_type=client_credentials' + `&client_id=${aadClientId}&client_secret=${aadClientSecret}`;

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/x-www-form-urlencoded',
		},
		content: requestBody,
	};

	const url = getMicrosoftTokenUrl(aadTenantId);
	const response = await http.post(url, httpRequest);

	if (response.statusCode === HttpStatusCode.OK) {
		const responseBody = response.data;
		if (responseBody === undefined) {
			throw new Error('Get application access token failed!');
		}

		const result: TokenResponse = {
			tokenType: responseBody.token_type,
			expiresIn: responseBody.expires_in,
			extExpiresIn: responseBody.ext_expires_in,
			accessToken: responseBody.access_token,
		};

		return result;
	}
	throw new Error(`Get application access token failed with http status code ${response.statusCode}.`);
};
