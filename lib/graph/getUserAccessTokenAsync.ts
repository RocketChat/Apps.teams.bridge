import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { BotUserAuthenticationScopes, getMicrosoftTokenUrl, NormalUserAuthenticationScopes } from "../Const";
import { TokenResponse } from './types';

export const getUserAccessTokenAsync = async (
    http: IHttp,
    accessCode: string,
    redirectUri: string,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string,
    userType: 'bot' | 'normal'): Promise<TokenResponse> => {
    const scopes = userType === 'bot' ? BotUserAuthenticationScopes : NormalUserAuthenticationScopes;
    const body = [
        `client_id=${encodeURIComponent(aadClientId)}`,
        `scope=${encodeURIComponent(scopes.join(' '))}`,
        `code=${encodeURIComponent(accessCode)}`,
        `redirect_uri=${encodeURIComponent(redirectUri)}`,
        `grant_type=authorization_code`,
        `client_secret=${encodeURIComponent(aadClientSecret)}`,
    ].join('&');

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        content: body
    };

    const url = getMicrosoftTokenUrl(aadTenantId);
    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get user access token failed!');
        }

        const result: TokenResponse = {
            tokenType: responseBody.token_type,
            expiresIn: responseBody.expires_in,
            extExpiresIn: responseBody.ext_expires_in,
            accessToken: responseBody.access_token,
            refreshToken: responseBody.refresh_token,
        };

        return result;
    } else {
        throw new Error(`Get user access token failed with http status code ${response.statusCode}.`);
    }
};
