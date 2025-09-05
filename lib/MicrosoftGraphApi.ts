import {
    HttpStatusCode,
    IHttp,
    IHttpRequest,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import * as ms from 'ms';
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
    SubscriptionMaxExpireTimeInSecond,
    getGraphApiUserUrl,
    getGraphApiChatMemberUrl,
    getGraphApiShareUrl,
    getGraphApiUploadToDriveUrl,
    getGraphApiChatMemberRemoveUrl,
    getGraphApiShareOneDriveFileUrl,
    getGraphApiOneDriveFileLinkUrl,
    getGraphApiRevokeRefreshTokenUrl,
    getGraphApiSubscriptionOperationUrl,
    getGraphApiChatThreadWithMemberUrl,
    RegistrationAutoRenewInterval,
} from "./Const";
import { getNotificationEndpointUrl } from "./UrlHelper";
import { getSubscriptionStateHashForUser } from "./PersistHelper";

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

export enum ThreadType {
    Group = 'group',
    OneOnOne = 'oneOnOne',
};

export interface GetThreadResponse {
    threadId: string;
    topic?: string;
    type?: ThreadType;
    memberIds?: string[];
};

export interface SendMessageResponse {
    messageId: string;
};

export interface SubscriptionResponse {
    subscriptionId: string;
    expirationTime: Date;
};

export enum MessageType {
    Message = 'message',
    SystemAddMembers = 'addMembers'
};

export enum MessageContentType {
    Html = 'html'
};

export interface Attachment {
    id: string,
    contentType: string,
    contentUrl: string | null,
    name: string | null,
};

export interface TeamsMessageReaction {
    reactionType: string; // e.g., "😆"
    displayName: string; // e.g., "Laugh"
    reactionContentUrl: string | null;
    createdDateTime: string; // ISO datetime string
    user: {
        application: string | null;
        device: string | null;
        user: {
            "@odata.type": "#microsoft.graph.teamworkUserIdentity";
            id: string;
            displayName: string | null;
            userIdentityType: "aadUser" | string;
            tenantId: string;
        };
    };
}
export interface GetMessageResponse {
    threadId: string;
    messageId: string;
    messageType: MessageType | undefined;
    fromUserTeamsId: string;
    messageContentType: MessageContentType | undefined;
    messageContent: string;
    attachments?: Attachment[];
    memberIds?: string[];
    reactions?: TeamsMessageReaction[];
};

export interface UploadFileResponse {
    driveItemId: string;
    fileName: string;
    size: number;
};

export interface ShareOneDriveFileResponse {
    shareId: string;
    shareLink: string;
};

export type SubscriptionsResponse = {
    "@odata.context": string
    value: Array<SubscriptionValue>
  }

type SubscriptionValue = {
    id: string
    resource: string
    applicationId: string
    changeType: string
    clientState: any
    notificationUrl: string
    notificationQueryOptions: any
    lifecycleNotificationUrl: any
    expirationDateTime: string
    creatorId: string
    includeResourceData: boolean
    latestSupportedTlsVersion: string
    encryptionCertificate: any
    encryptionCertificateId: any
    notificationUrlAppId: any
}

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
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get application access token failed!');
        }

        const result : TokenResponse = {
            tokenType: responseBody.token_type,
            expiresIn: responseBody.expires_in,
            extExpiresIn: responseBody.ext_expires_in,
            accessToken: responseBody.access_token,
        };

        return result;
    } else {
        throw new Error(`Get application access token failed with http status code ${response.statusCode}.`);
    }
};

export const listTeamsUserProfilesAsync = async (
    http: IHttp,
    appAccessToken: string) : Promise<TeamsUserProfile[]> => {
    const url = getGraphApiUserUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${appAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('List users failed!');
        }


        const userList = responseBody.value as any[];
        const result : TeamsUserProfile[] = [];
        for (let index = 0; index < userList.length; index++) {
            try {
                const user = userList[index];
                const record : TeamsUserProfile = {
                    displayName: user.displayName,
                    givenName: user.givenName,
                    surname: user.surname,
                    mail: user.mail,
                    id: user.id,
                };

                result.push(record);
            } catch (error) {
                // If there's an error, print a warning but not block the whole process
                console.error(`Error when handling user list. Details: ${error}`);
            }
        }

        return result;
    } else {
        throw new Error(`List users failed with http status code ${response.statusCode}.`);
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
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get user access token failed!');
        }


        const result : TokenResponse = {
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

export const renewUserAccessTokenAsync = async (
    http: IHttp,
    refreshToken: string,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string) : Promise<TokenResponse> => {
    let body = `client_id=${aadClientId}`;
    body += `&scope=${AuthenticationScopes.join(' ')}`;
    body += `&refresh_token=${refreshToken}`;
    body += `&grant_type=refresh_token`;
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
            throw new Error('Refresh user access token failed!');
        }


        const result : TokenResponse = {
            tokenType: responseBody.token_type,
            expiresIn: responseBody.expires_in,
            extExpiresIn: responseBody.ext_expires_in,
            accessToken: responseBody.access_token,
            refreshToken: responseBody.refresh_token,
        };

        return result;
    } else {
        throw new Error(`Refresh user access token failed with http status code ${response.statusCode}.`);
    }
};

export const revokeUserRefreshTokenAsync = async (http: IHttp, userAccessToken: string) : Promise<void> => {
    const url = getGraphApiRevokeRefreshTokenUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json',
        },
        content: '{}'
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        return;
    } else {
        throw new Error(`Revoke user refresh token failed with http status code ${response.statusCode}.\nReceived: ${JSON.stringify(response.data, null, 2)}`);
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
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get user profile failed!');
        }


        const result : TeamsUserProfile = {
            displayName: responseBody.displayName,
            givenName: responseBody.givenName,
            surname: responseBody.surname,
            mail: responseBody.mail,
            id: responseBody.id,
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
        chatType: 'oneOnOne',
        members: [senderUserTeamsId, receiverUserTeamsId].map(userId => ({
            '@odata.type': '#microsoft.graph.aadUserConversationMember',
            roles: ['owner'],
            'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`,
        })),
    };

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body)
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Create one on one chat thread failed!');
        }


        const result : CreateThreadResponse = {
            threadId: responseBody.id,
        };

        return result;
    } else {
        throw new Error(`Create one on one chat thread failed with http status code ${response.statusCode}.\n Received: ${JSON.stringify(response.data, null, 2)}`);
    }
};

export const createChatThreadAsync = async (
    http: IHttp,
    membersTeamsIds: string[],
    roomName: string,
    userAccessToken: string) : Promise<CreateThreadResponse> => {
    const url = getGraphApiChatUrl();

    const uniqueMembersTeamsIds = [...new Set(membersTeamsIds)];

    const members: any[] = [];
    for (const teamsIds of uniqueMembersTeamsIds) {
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
        'topic': roomName,
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
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Create group chat thread failed!');
        }


        const result : CreateThreadResponse = {
            threadId: responseBody.id,
        };

        return result;
    } else {
        throw new Error(
            `Create group chat thread failed with http status code ${
                response.statusCode
            }.\nReceived: ${JSON.stringify(response.data, null, 2)}`
        );
    }
};

export const getChatThreadWithMembersAsync = async (
    http: IHttp,
    threadId: string,
    userAccessToken: string) : Promise<GetThreadResponse> => {
    const url = getGraphApiChatThreadWithMemberUrl(threadId);

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get chat thread failed!');
        }



        let memberIds : string[] | undefined = undefined;

        const jsonMembers = responseBody.members as any[];
        if (jsonMembers) {
            memberIds = [];
            for (const jsonMember of jsonMembers) {
                memberIds.push(jsonMember.userId);
            }
        }

        const result : GetThreadResponse = {
            threadId: responseBody.id,
            topic: responseBody.topic,
            type: parseThreadType(responseBody.chatType),
            memberIds: memberIds,
        };

        return result;
    } else {
        throw new Error(`Get chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const addMemberToChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    memberTeamsIdToBeAdd: string,
    userAccessToken: string) : Promise<void> => {
    const url = getGraphApiChatMemberUrl(threadId);

    const body = {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        'roles': ['owner'],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${memberTeamsIdToBeAdd}')`,
        'visibleHistoryStartDateTime': '0001-01-01T00:00:00Z',
    };

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body)
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode !== HttpStatusCode.CREATED) {
        throw new Error(`Add member to group chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const listMembersInChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    userAccessToken: string) : Promise<string[]> => {
    const url = getGraphApiChatMemberUrl(threadId);
    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('List members in chat thread failed!');
        }


        const userList = responseBody.value as any[];

        const result : string[] = [];

        for (const user of userList) {
            result.push(user.userId);
        }

        return result;
    } else {
        throw new Error(`List members in chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const removeMemberFromChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    teamsUserId: string,
    userAccessToken: string) : Promise<void> => {
    const url = getGraphApiChatMemberRemoveUrl(threadId, teamsUserId);
    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.del(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        return;
    } else {
        throw new Error(`Remove member from chat thread failed with http status code ${response.statusCode}.`);
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
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Send message to chat thread failed!');
        }


        const result : SendMessageResponse = {
            messageId: responseBody.id,
        };

        return result;
    } else {
        throw new Error(`Send message to chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const sendFileMessageToChatThreadAsync = async (
    http: IHttp,
    textMessage: string,
    fileName: string,
    shareLink: string,
    threadId: string,
    userAccessToken: string) : Promise<SendMessageResponse> => {
    const url = getGraphApiMessageUrl(threadId);

    const body = {
        'body' : {
            'content': `${textMessage} <a href=\"${shareLink}\" title=\"${shareLink}\" target=\"_blank\" rel=\"noreferrer noopener\">${fileName}</a>`,
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
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Send file message to chat thread failed!');
        }


        const result : SendMessageResponse = {
            messageId: responseBody.id,
        };

        return result;
    } else {
        throw new Error(`Send file message to chat thread failed with http status code ${response.statusCode}.`);
    }
};

export const updateTextMessageInChatThreadAsync = async (
    http: IHttp,
    textMessage: string,
    messageType: 'text' | 'html',
    messageId: string,
    threadId: string,
    userAccessToken: string) : Promise<void> => {
    const url = getGraphApiMessageUrl(threadId, messageId, true);

    const body = {
        'body' : {
            'content': textMessage,
            'contentType': messageType
        },
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
        return;
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
        return;
    } else {
        console.log(`Response Error: ${response.content}`)
        // throw new Error(`Delete message in chat thread failed with http status code ${response.statusCode}.`);
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
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get message with resource string failed!');
        }

        let attachments : Attachment[] | undefined = undefined;

        const jsonAttachments = responseBody.attachments as any[];
        if (jsonAttachments && jsonAttachments.length > 0) {
            attachments = [];
            for (const jsonAttachment of jsonAttachments) {
                const attachment : Attachment = {
                    id: jsonAttachment.id,
                    contentType: jsonAttachment.contentType,
                    contentUrl: jsonAttachment.contentUrl,
                    name: jsonAttachment.name,
                };
                attachments.push(attachment);
            }
        }

        const messageType = parseMessageType(responseBody.messageType, responseBody.eventDetail);

        let memberIds : string[] | undefined = undefined;
        if (messageType === MessageType.SystemAddMembers) {
            memberIds = [];
            const jsonMembers = responseBody.eventDetail.members as any[];
            for (const jsonMember of jsonMembers) {
                memberIds.push(jsonMember.id);
            }
        }

        const result : GetMessageResponse = {
            threadId: responseBody.chatId,
            messageId: responseBody.id,
            messageType: messageType,
            fromUserTeamsId: responseBody.from?.user?.id,
            messageContentType: parseMessageContentType(responseBody.body?.contentType),
            messageContent: responseBody.body?.content,
            attachments: attachments,
            memberIds: memberIds,
            reactions: responseBody.reactions,
        };

        return result;
    } else {
        throw new Error(`Get message with resource string failed with http status code ${response.statusCode}.`);
    }
}

export const listSubscriptionsAsync = async (
    http: IHttp,
    userAccessToken: string,
    notificationUrl: string,
): Promise<SubscriptionsResponse['value'] | undefined> => {
    const url = getGraphApiSubscriptionUrl();

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error("List subscriptions failed!");
        }

        const subscriptions = response.data?.value as SubscriptionsResponse['value'];
        const urlObj = new URL(notificationUrl);
        const pathWithQuery = urlObj.pathname + urlObj.search;
        return subscriptions.filter((subscription) => subscription.notificationUrl.includes(pathWithQuery));
    } else {
        console.error(
            `List subscriptions failed with http status code ${response.statusCode}. \nReceived: ${JSON.stringify(response.data, null, 2)}`
        );
        return;
    }
};

export const renewSubscriptionAsync = async (
    http: IHttp,
    subscriptionId: string,
    userAccessToken: string,
    expirationDateTime?: Date,
    clientState?: string
): Promise<SubscriptionResponse | undefined> => {
    if (!expirationDateTime) {
        expirationDateTime = new Date();
        expirationDateTime.setSeconds(
            expirationDateTime.getSeconds() + SubscriptionMaxExpireTimeInSecond
        );
    }

    const url = getGraphApiSubscriptionOperationUrl(subscriptionId);

    const body = {
        expirationDateTime: expirationDateTime.toISOString(),
        ...(clientState && { clientState }),
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.patch(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error("Renew subscription failed!");
        }

        const result: SubscriptionResponse = {
            subscriptionId: responseBody.id,
            expirationTime: new Date(responseBody.expirationDateTime),
        };

        return result;
    } else {
        console.error(
            `Renew subscription failed with http status code ${
                response.statusCode
            }.\nReceived: ${JSON.stringify(response.data, null, 2)}`
        );
        return;
    }
};

export const deleteAllSubscriptions = async (http: IHttp, userAccessToken: string, notificationUrl: string) => {
    // Delete all subscriptions
    const subscriptionsId = (
        await listSubscriptionsAsync(
            http,
            userAccessToken,
            notificationUrl,
        )
    )?.map((subscription) => (subscription as any).id);
    if (subscriptionsId) {
        for (const subscriptionId of subscriptionsId) {
            try {
                await deleteSubscriptionAsync(
                    http,
                    subscriptionId,
                    userAccessToken
                );
            } catch (error) {
                console.error(
                    `Error during delete subscription, will ignore and continue. ${error}`
                );
            }
        }
    }
}

export const deleteSubscriptionAsync = async (
    http: IHttp,
    subscriptionId: string,
    userAccessToken: string) : Promise<void> => {
    const url = getGraphApiSubscriptionOperationUrl(subscriptionId);

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.del(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        return;
    } else {
        throw new Error(`Delete subscription failed with http status code ${response.statusCode}.`);
    }
};

export const subscribeToAllMessagesForOneUserAsync = async (options: {
    http: IHttp;
    read: IRead;
    persis: IPersistence,
    rocketChatUserId: string;
    teamsUserId: string;
    subscriberEndpointUrl: string;
    userAccessToken: string;
    expirationDateTime?: Date;
    renewIfExists?: boolean;
    forceRenew?: boolean;
}): Promise<SubscriptionResponse | undefined> => {
    const {
        rocketChatUserId,
        teamsUserId,
        subscriberEndpointUrl,
        userAccessToken,
        expirationDateTime: inputExpirationDateTime,
        renewIfExists = true,
        forceRenew = false,
        http,
        read,
        persis,
    } = options;

    let expirationDateTime = inputExpirationDateTime;
    if (!expirationDateTime) {
        expirationDateTime = new Date();
        expirationDateTime.setSeconds(
            expirationDateTime.getSeconds() + SubscriptionMaxExpireTimeInSecond
        );
    }

    const url = getGraphApiSubscriptionUrl();

    const notificationUrl = getNotificationEndpointUrl({
        subscriberEndpoint: subscriberEndpointUrl,
        rocketChatUserId,
    });

    const clientState = await getSubscriptionStateHashForUser(
        read.getPersistenceReader(),
        persis,
        { rocketChatUserId }
    );
    const body = {
        changeType: SupportedNotificationChangeTypes.join(","),
        notificationUrl,
        resource: `/users/${teamsUserId}/chats/getAllMessages`,
        includeResourceData: false,
        expirationDateTime: expirationDateTime.toISOString(),
        clientState,
    };

    if (renewIfExists) {
        const existingSubscriptions = await listSubscriptionsAsync(http, userAccessToken, notificationUrl) || [];

        if (existingSubscriptions.length > 0) {
            if (existingSubscriptions.length > 1) {
                await Promise.all(
                    existingSubscriptions
                        .slice(1)
                        .map((sub) =>
                            deleteSubscriptionAsync(
                                http,
                                sub.id,
                                userAccessToken
                            )
                        )
                );
            }
            // For older versions of the app, clientState is not set
            const hasClientState = new URL(
                existingSubscriptions[0].notificationUrl
            ).searchParams.has('hasClientState');
            if (!hasClientState) {
                await deleteSubscriptionAsync(
                    http,
                    existingSubscriptions[0].id,
                    userAccessToken
                );
            } else {
                const existingSub = existingSubscriptions[0];
                const currentExpireTime = new Date(existingSub.expirationDateTime);
                const now = new Date();
                const nextUpdateTime = new Date(
                    now.getTime() + ms(RegistrationAutoRenewInterval)
                );

                const timeLeftAtNextUpdate = currentExpireTime.getTime() - nextUpdateTime.getTime();

                const threshold = ms(RegistrationAutoRenewInterval) / 2;

                const shouldRenew = timeLeftAtNextUpdate <= threshold;

                if (!shouldRenew && !forceRenew) {
                    return;
                }
                return await renewSubscriptionAsync(
                    http,
                    existingSubscriptions[0].id,
                    userAccessToken,
                    expirationDateTime,
                    clientState
                );
            }
        }
    }

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error("Subscribe to notification for user failed!");
        }

        const result: SubscriptionResponse = {
            subscriptionId: responseBody.id,
            expirationTime: new Date(responseBody.expirationDateTime),
        };

        return result;
    } else {
        throw new Error(
            `Subscribe to notification for user failed with http status code ${response.statusCode}.\nReceived: ${JSON.stringify(
                response.data,
                null,
                2
            )}`
        );
    }
};

export const downloadOneDriveFileAsync = async(
    http: IHttp,
    encodedUrl: string,
    userAccessToken: string) : Promise<Buffer> => {
    const url = getGraphApiShareUrl(encodedUrl);

    const httpRequest: IHttpRequest = {
        headers: {
            'Authorization': `Bearer ${userAccessToken}`,
        },
        encoding: null
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const fileStr = response.content as string;
        const buff = Buffer.from(fileStr, 'binary');
        return buff;
    } else {
        throw new Error(`Download one drive file failed with http status code ${response.statusCode}.`);
    }
};

export const uploadFileToOneDriveAsync = async(
    http: IHttp,
    fileName: string,
    fileMIMEType: string,
    fileSize: number,
    content: Buffer,
    userAccessToken: string) : Promise<UploadFileResponse | undefined> => {
    if (fileSize < 4096000) {
        // use direct upload
        const url = getGraphApiUploadToDriveUrl(fileName);

        const httpRequest: IHttpRequest = {
            headers: {
                'Authorization': `Bearer ${userAccessToken}`,
                'Content-Type': fileMIMEType,
            },
            content: content as any as string
        };

        const response = await http.put(url, httpRequest);

        if ([HttpStatusCode.CREATED, HttpStatusCode.OK].includes(response.statusCode)) {
            const responseBody = response.data;
            if (responseBody === undefined) {
                throw new Error('Upload file to one drive failed!');
            }


            const result : UploadFileResponse = {
                driveItemId: responseBody.id,
                fileName: responseBody.name,
                size: responseBody.size,
            };

            return result;
        } else {
            throw new Error(`Upload file to one drive failed with http status code ${response.statusCode}.`);
        }
    } else {
        // TODO: implement resumable upload
        return undefined;
    }
};

export const shareOneDriveFileAsync = async (
    http: IHttp,
    oneDriveItemId: string,
    userAccessToken: string) : Promise<ShareOneDriveFileResponse> => {
    const url = getGraphApiShareOneDriveFileUrl(oneDriveItemId);

    const body = {
        'type': 'view',
        'scope': 'organization',
    };

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if ([HttpStatusCode.CREATED, HttpStatusCode.OK].includes(response.statusCode)) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Create share link for onedrive item failed!');
        }


        const result : ShareOneDriveFileResponse = {
            shareId: responseBody.id,
            shareLink: responseBody.link.webUrl,
        };

        return result;
    } else {
        throw new Error(`Create share link for onedrive item failed with http status code ${response.statusCode}.`);
    }
};

export const getOneDriveFileLinkAsync = async (
    http: IHttp,
    oneDriveItemId: string,
    userAccessToken: string): Promise<string> => {
    const url = getGraphApiOneDriveFileLinkUrl(oneDriveItemId);

    const httpRequest: IHttpRequest = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.data;
        if (responseBody === undefined) {
            throw new Error('Get one drive file link failed!');
        }



        const result : string = responseBody.webUrl as string;

        return result;
    } else {
        throw new Error(`Get one drive file link failed with http status code ${response.statusCode}.`);
    }
};

const parseMessageType = (messageType: string, eventDetail?: any) : MessageType | undefined => {
    if (!messageType) {
        return undefined;
    }

    if (messageType === 'message') {
        return MessageType.Message;
    } else {
        if (eventDetail) {
            if (eventDetail['@odata.type'] === '#microsoft.graph.membersAddedEventMessageDetail') {
                return MessageType.SystemAddMembers;
            }
        }
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

const parseThreadType = (threadType: string) : ThreadType | undefined => {
    if (!threadType) {
        return undefined;
    }

    if (threadType === 'group') {
        return ThreadType.Group;
    }

    if (threadType === 'oneOnOne') {
        return ThreadType.OneOnOne;
    }

    return undefined;
};
