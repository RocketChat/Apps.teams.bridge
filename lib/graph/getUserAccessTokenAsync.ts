import { HttpStatusCode, IHttp, IHttpRequest } from "@rocket.chat/apps-engine/definition/accessors";
import { AuthenticationScopes, getMicrosoftTokenUrl } from "../Const";
import { TokenResponse } from './types';

export const getUserAccessTokenAsync = async (
    http: IHttp,
    accessCode: string,
    redirectUri: string,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string): Promise<TokenResponse> => {
    let body = `client_id=${aadClientId}`;
    body += `&scope=${AuthenticationScopes.join(' ')}`;
    body += `&code=${accessCode}`;
    body += `&redirect_uri=${redirectUri}`;
    body += `&grant_type=authorization_code`;
    body += `&client_secret=${aadClientSecret}`;
    body = encodeURI(body);

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
