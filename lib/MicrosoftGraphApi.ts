import {
    HttpStatusCode,
    IHttp,
    IHttpRequest,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    AuthenticationScopes,
    getGraphApiChatUrl,
    getGraphApiMessageDeleteUrl,
    getGraphApiMessageUrl,
    getGraphApiProfileUrl,
    getGraphApiResourceUrl,
    getGraphApiSubscriptionUrl,
    getMicrosoftTokenUrl,
    SupportedNotificationChangeTypes,
    SubscriptionMaxExpireTimeInSecond
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

export interface CreateSubscriptionResponse {
    subscriptionId: string;
    expirationTime: Date;
};

export enum MessageType {
    Message = 'message'
};

export enum MessageContentType {
    Html = 'html'
};

export interface GetMessageResponse {
    threadId: string;
    messageId: string;
    messageType: MessageType | undefined;
    fromUserTeamsId: string;
    messageContentType: MessageContentType | undefined;
    messageContent: string;
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

export const createChatThreadAsync = async (
    http: IHttp,
    membersTeamsIds: string[],
    bridgeUserName: string,
    userAccessToken: string) : Promise<CreateThreadResponse> => {
    const url = getGraphApiChatUrl();

    const members: any[] = [];
    for (const teamsIds of membersTeamsIds) {
        const member = {
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            'roles': ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${teamsIds}')`,
        };
        members.push(member);
    }

    const body = {
        'chatType': 'group',
        'members': members,
        'topic': `Rocket.Chat interop bridged by ${bridgeUserName}`,
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
            throw new Error('Create group chat thread failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : CreateThreadResponse = {
            threadId: jsonBody.id,
        };

        console.log('teams group thread created!');
        console.log(result);

        return result;
    } else {
        throw new Error(`Create group chat thread failed with http status code ${response.statusCode}.`);
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
            'content': textMessage,
            'contentType': 'html'
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

export const updateTextMessageInChatThreadAsync = async (
    http: IHttp,
    textMessage: string,
    messageId: string,
    threadId: string,
    userAccessToken: string) : Promise<void> => {
    const url = getGraphApiMessageUrl(threadId, messageId, true);

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

    const response = await http.patch(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Update message in chat thread failed!');
        }
    } else {
        throw new Error(`Update message in chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const deleteTextMessageInChatThreadAsync = async (
    http: IHttp,
    teamsUserId: string,
    messageId: string,
    threadId: string,
    userAccessToken: string) : Promise<void> => {
    const url = getGraphApiMessageDeleteUrl(teamsUserId, threadId, messageId);

    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Delete message in chat thread failed!');
        }
    } else {
        throw new Error(`Delete message in chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const getMessageWithResourceStringAsync = async (
    http: IHttp,
    resourceString: string,
    userAccessToken: string) : Promise<GetMessageResponse> => {

    const url = getGraphApiResourceUrl(resourceString);
    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error('Get message with resource string failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : GetMessageResponse = {
            threadId: jsonBody.chatId,
            messageId: jsonBody.id,
            messageType: parseMessageType(jsonBody.messageType),
            fromUserTeamsId: jsonBody.from?.user?.id,
            messageContentType: parseMessageContentType(jsonBody.body?.contentType),
            messageContent: jsonBody.body?.content,
        };

        return result;
    } else {
        throw new Error(`Get message with resource string failed with http status code ${response.statusCode}.`);
    }
}

export const subscribeToAllMessagesForOneUserAsync = async (
    http: IHttp,
    rocketChatUserId: string,
    teamsUserId: string,
    subscriberEndpointUrl: string,
    userAccessToken: string,
    expirationDateTime?: Date
) : Promise<CreateSubscriptionResponse> => {

    if (!expirationDateTime) {
        expirationDateTime = new Date();
        expirationDateTime.setSeconds(expirationDateTime.getSeconds() + SubscriptionMaxExpireTimeInSecond);
    }

    const url = getGraphApiSubscriptionUrl();

    const body = {
        'changeType': SupportedNotificationChangeTypes.join(','),
        'notificationUrl': `${subscriberEndpointUrl}?userId=${rocketChatUserId}`,
        'resource': `/users/${teamsUserId}/chats/getAllMessages`,
        'includeResourceData': false,
        'expirationDateTime': expirationDateTime.toISOString()
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
            throw new Error('Subscribe to notification for user failed!');
        }

        const jsonBody = JSON.parse(responseBody);
        const result : CreateSubscriptionResponse = {
            subscriptionId: jsonBody.id,
            expirationTime: new Date(jsonBody.expirationDateTime),
        };
    
        console.log("VVV==CreateSubscriptionResponse==VVV");
        console.log(result);
        console.log("^^^==CreateSubscriptionResponse==^^^");
    
        return result;
    } else {
        throw new Error(`Subscribe to notification for user failed with http status code ${response.statusCode}.`);
    }
};

const parseMessageType = (messageType: string) : MessageType | undefined => {
    if (!messageType) {
        return undefined;
    }

    if (messageType === 'message') {
        return MessageType.Message;
    }

    return undefined;
};

const parseMessageContentType = (messageContentType: string) : MessageContentType | undefined => {
    if (!messageContentType) {
        return undefined;
    }

    if (messageContentType === 'html') {
        return MessageContentType.Html;
    }

    return undefined;
};
