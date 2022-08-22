import {
    HttpStatusCode,
    IHttp,
    IHttpRequest,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    AuthenticationScopes,
    getGraphApiChatUrl,
    getGraphApiMessageUrl,
    getGraphApiProfileUrl,
    getMicrosoftTokenUrl
} from "./Const";

export interface TokenResponse {
    tokenType: string;
    expiresIn: number;
    extExpiresIn: number;
    accessToken: string;
    refreshToken?: string;
};

export interface TeamsUserProfile {
    displayName: string;
    givenName: string;
    surname: string;
    mail: string;
    id: string;
};

export interface CreateThreadResponse {
    threadId: string;
};

export interface SendMessageResponse {
    messageId: string;
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
            'Content-Type': 'application/x-www-form-urlencoded'
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

export const getUserProfileAsync = async (http: IHttp, userAccessToken: string) : Promise<TeamsUserProfile> => {
    const url = getGraphApiProfileUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Get user profile failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : TeamsUserProfile = {
            displayName: jsonBody.displayName,
            givenName: jsonBody.givenName,
            surname: jsonBody.surname,
            mail: jsonBody.mail,
            id: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(`Get user profile failed with http status code ${response.statusCode}.`);
    }
};

export const createOneOnOneChatThreadAsync = async (
    http: IHttp,
    senderUserTeamsId: string,
    receiverUserTeamsId: string,
    userAccessToken: string) : Promise<CreateThreadResponse> => {
    const url = getGraphApiChatUrl();

    const body = {
        'chatType': 'oneOnOne',
        'members': [
            {
                '@odata.type': '#microsoft.graph.aadUserConversationMember',
                'roles': ['owner'],
                'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${senderUserTeamsId}')`,
            },
            {
                '@odata.type': '#microsoft.graph.aadUserConversationMember',
                'roles': ['owner'],
                'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${receiverUserTeamsId}')`,
            },
        ],
    }

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body)
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Create one on one chat thread failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : CreateThreadResponse = {
            threadId: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(`Create one on one chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const sendTextMessageToChatThreadAsync = async (
    http: IHttp,
    textMessage: string,
    threadId: string,
    userAccessToken: string) : Promise<SendMessageResponse> => {
    const url = getGraphApiMessageUrl(threadId);

    const body = {
        'body' : {
            'content': textMessage
        }
    }

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body)
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Send message to chat thread failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : SendMessageResponse = {
            messageId: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(`Send message to chat thread failed with http status code ${response.statusCode}.`);
    }
};
