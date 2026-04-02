import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { BotUserAuthenticationScopes, getMicrosoftTokenUrl, NormalUserAuthenticationScopes } from '../Const';
import type { TokenResponse } from './types';

export const renewUserAccessTokenAsync = async (
	http: IHttp,
	refreshToken: string,
	aadTenantId: string,
	aadClientId: string,
	aadClientSecret: string,
	type: 'bot' | 'normal',
): Promise<TokenResponse> => {
	const scopes = type === 'bot' ? BotUserAuthenticationScopes : NormalUserAuthenticationScopes;
	const body = [
		`client_id=${encodeURIComponent(aadClientId)}`,
		`scope=${encodeURIComponent(scopes.join(' '))}`,
		`refresh_token=${encodeURIComponent(refreshToken)}`,
		`grant_type=refresh_token`,
		`client_secret=${encodeURIComponent(aadClientSecret)}`,
	].join('&');

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/x-www-form-urlencoded',
		},
		content: body,
	};

	const url = getMicrosoftTokenUrl(aadTenantId);
	const response = await http.post(url, httpRequest);

	if (response.statusCode === HttpStatusCode.OK) {
		const responseBody = response.data;
		if (responseBody === undefined) {
			throw new Error('Refresh user access token failed!');
		}

		const result: TokenResponse = {
			tokenType: responseBody.token_type,
			expiresIn: responseBody.expires_in,
			extExpiresIn: responseBody.ext_expires_in,
			accessToken: responseBody.access_token,
			refreshToken: responseBody.refresh_token,
		};

		return result;
	}
	throw new Error(`Refresh user access token failed with http status code ${response.statusCode}.`);
};
