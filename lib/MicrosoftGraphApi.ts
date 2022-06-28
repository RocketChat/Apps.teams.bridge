import {
    HttpStatusCode,
    IHttp,
    IHttpRequest,
} from "@rocket.chat/apps-engine/definition/accessors";
import { AuthenticationScopes, getMicrosoftTokenUrl } from "./Const";

export interface TokenResponse {
    tokenType: string;
    expiresIn: number;
    extExpiresIn: number;
    accessToken: string;
    refreshToken?: string;
};

export const getApplicationAccessTokenAsync = async (
    http: IHttp,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string) : Promise<TokenResponse> => {
    const requestBody = 'scope=https://graph.microsoft.com/.default&grant_type=client_credentials'
        + `&client_id=${aadClientId}&client_secret=${aadClientSecret}`;

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        content: requestBody,
    };

    const url = getMicrosoftTokenUrl(aadTenantId);
    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Get application access token failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : TokenResponse = {
            tokenType: jsonBody.token_type,
            expiresIn: jsonBody.expires_in,
            extExpiresIn: jsonBody.ext_expires_in,
            accessToken: jsonBody.access_token,
        };

        return result;
    } else {
        throw new Error(`Get application access token failed with http status code ${response.statusCode}.`);
    }
};

export const getUserAccessTokenAsync = async (
    http: IHttp,
    accessCode: string,
    redirectUri: string,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string) : Promise<TokenResponse> => {
    let body = `client_id=${aadClientId}`;
    body += `&scope=${AuthenticationScopes.join(' ')}`;
    body += `&code=${accessCode}`;
    body += `&redirect_uri=${redirectUri}`;
    body += `&grant_type=authorization_code`;
    body += `&client_secret=${aadClientSecret}`;
    body = encodeURI(body);
    
    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/x-www-form-urlencoded"
        },
        content: body
    };

    const url = getMicrosoftTokenUrl(aadTenantId);
    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Get application access token failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : TokenResponse = {
            tokenType: jsonBody.token_type,
            expiresIn: jsonBody.expires_in,
            extExpiresIn: jsonBody.ext_expires_in,
            accessToken: jsonBody.access_token,
            refreshToken: jsonBody.refresh_token,
        };

        return result;
    } else {
        throw new Error(`Get application access token failed with http status code ${response.statusCode}.`);
    }
};
