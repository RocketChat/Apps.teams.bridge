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
} from "./Const";

export interface TokenResponse {
    tokenType: string;
    expiresIn: number;
    extExpiresIn: number;
    accessToken: string;
    refreshToken?: string;
}

export interface TeamsUserProfile {
    displayName: string;
    givenName: string;
    surname: string;
    mail: string;
    id: string;
}

export interface CreateThreadResponse {
    threadId: string;
}

export enum ThreadType {
    Group = "group",
    OneOnOne = "oneOnOne",
}

export interface GetThreadResponse {
    threadId: string;
    topic?: string;
    type?: ThreadType;
    memberIds?: string[];
}

export interface SendMessageResponse {
    messageId: string;
}

export interface SubscriptionResponse {
    subscriptionId: string;
    expirationTime: Date;
}

export enum MessageType {
    Message = "message",
    SystemAddMembers = "addMembers",
}

export enum MessageContentType {
    Html = "html",
}

export interface Attachment {
    id: string;
    contentType: string;
    contentUrl: string;
    name: string;
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
}

export interface UploadFileResponse {
    driveItemId: string;
    fileName: string;
    size: number;
}

export interface ShareOneDriveFileResponse {
    shareId: string;
    shareLink: string;
}

export const getApplicationAccessTokenAsync = async (
    http: IHttp,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string
): Promise<TokenResponse> => {
    const requestBody =
        "scope=https://graph.microsoft.com/.default&grant_type=client_credentials" +
        `&client_id=${aadClientId}&client_secret=${aadClientSecret}`;

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/x-www-form-urlencoded",
        },
        content: requestBody,
    };

    const url = getMicrosoftTokenUrl(aadTenantId);
    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Get application access token failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: TokenResponse = {
            tokenType: jsonBody.token_type,
            expiresIn: jsonBody.expires_in,
            extExpiresIn: jsonBody.ext_expires_in,
            accessToken: jsonBody.access_token,
        };

        return result;
    } else {
        throw new Error(
            `Get application access token failed with http status code ${response.statusCode}.`
        );
    }
};

export const listTeamsUserProfilesAsync = async (
    http: IHttp,
    appAccessToken: string
): Promise<TeamsUserProfile[]> => {
    const url = getGraphApiUserUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            Authorization: `Bearer ${appAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("List users failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const userList = jsonBody.value as any[];
        const result: TeamsUserProfile[] = [];
        for (let index = 0; index < userList.length; index++) {
            try {
                const user = userList[index];
                const record: TeamsUserProfile = {
                    displayName: user.displayName,
                    givenName: user.givenName,
                    surname: user.surname,
                    mail: user.mail,
                    id: user.id,
                };

                result.push(record);
            } catch (error) {
                // If there's an error, print a warning but not block the whole process
                console.error(
                    `Error when handling user list. Details: ${error}`
                );
            }
        }

        return result;
    } else {
        throw new Error(
            `List users failed with http status code ${response.statusCode}.`
        );
    }
};

export const getUserAccessTokenAsync = async (
    http: IHttp,
    accessCode: string,
    redirectUri: string,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string
): Promise<TokenResponse> => {
    let body = `client_id=${aadClientId}`;
    body += `&scope=${AuthenticationScopes.join(" ")}`;
    body += `&code=${accessCode}`;
    body += `&redirect_uri=${redirectUri}`;
    body += `&grant_type=authorization_code`;
    body += `&client_secret=${aadClientSecret}`;
    body = encodeURI(body);

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/x-www-form-urlencoded",
        },
        content: body,
    };

    const url = getMicrosoftTokenUrl(aadTenantId);
    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Get user access token failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: TokenResponse = {
            tokenType: jsonBody.token_type,
            expiresIn: jsonBody.expires_in,
            extExpiresIn: jsonBody.ext_expires_in,
            accessToken: jsonBody.access_token,
            refreshToken: jsonBody.refresh_token,
        };

        return result;
    } else {
        throw new Error(
            `Get user access token failed with http status code ${response.statusCode}.`
        );
    }
};

export const renewUserAccessTokenAsync = async (
    http: IHttp,
    refreshToken: string,
    aadTenantId: string,
    aadClientId: string,
    aadClientSecret: string
): Promise<TokenResponse> => {
    let body = `client_id=${aadClientId}`;
    body += `&scope=${AuthenticationScopes.join(" ")}`;
    body += `&refresh_token=${refreshToken}`;
    body += `&grant_type=refresh_token`;
    body += `&client_secret=${aadClientSecret}`;
    body = encodeURI(body);

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/x-www-form-urlencoded",
        },
        content: body,
    };

    const url = getMicrosoftTokenUrl(aadTenantId);
    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Refresh user access token failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: TokenResponse = {
            tokenType: jsonBody.token_type,
            expiresIn: jsonBody.expires_in,
            extExpiresIn: jsonBody.ext_expires_in,
            accessToken: jsonBody.access_token,
            refreshToken: jsonBody.refresh_token,
        };

        return result;
    } else {
        throw new Error(
            `Refresh user access token failed with http status code ${response.statusCode}.`
        );
    }
};

export const revokeUserRefreshTokenAsync = async (
    http: IHttp,
    userAccessToken: string
): Promise<void> => {
    const url = getGraphApiRevokeRefreshTokenUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        return;
    } else {
        throw new Error(
            `Revoke user refresh token failed with http status code ${response.statusCode}.`
        );
    }
};

export const getUserProfileAsync = async (
    http: IHttp,
    userAccessToken: string
): Promise<TeamsUserProfile> => {
    const url = getGraphApiProfileUrl();
    const httpRequest: IHttpRequest = {
        headers: {
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Get user profile failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: TeamsUserProfile = {
            displayName: jsonBody.displayName,
            givenName: jsonBody.givenName,
            surname: jsonBody.surname,
            mail: jsonBody.mail,
            id: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(
            `Get user profile failed with http status code ${response.statusCode}.`
        );
    }
};

export const createOneOnOneChatThreadAsync = async (
    http: IHttp,
    senderUserTeamsId: string,
    receiverUserTeamsId: string,
    userAccessToken: string
): Promise<CreateThreadResponse> => {
    const url = getGraphApiChatUrl();

    const body = {
        chatType: "oneOnOne",
        members: [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                roles: ["owner"],
                "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${senderUserTeamsId}')`,
            },
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                roles: ["owner"],
                "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${receiverUserTeamsId}')`,
            },
        ],
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Create one on one chat thread failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: CreateThreadResponse = {
            threadId: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(
            `Create one on one chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const createChatThreadAsync = async (
    http: IHttp,
    membersTeamsIds: string[],
    roomName: string,
    userAccessToken: string
): Promise<CreateThreadResponse> => {
    const url = getGraphApiChatUrl();

    const members: any[] = [];
    for (const teamsIds of membersTeamsIds) {
        const member = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            roles: ["owner"],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${teamsIds}')`,
        };
        members.push(member);
    }

    const body = {
        chatType: "group",
        members: members,
        topic: roomName,
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Create group chat thread failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: CreateThreadResponse = {
            threadId: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(
            `Create group chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const getChatThreadWithMembersAsync = async (
    http: IHttp,
    threadId: string,
    userAccessToken: string
): Promise<GetThreadResponse> => {
    const url = getGraphApiChatThreadWithMemberUrl(threadId);

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Get chat thread failed!");
        }

        const jsonBody = JSON.parse(responseBody);

        let memberIds: string[] | undefined = undefined;

        const jsonMembers = jsonBody.members as any[];
        if (jsonMembers) {
            memberIds = [];
            for (const jsonMember of jsonMembers) {
                memberIds.push(jsonMember.userId);
            }
        }

        const result: GetThreadResponse = {
            threadId: jsonBody.id,
            topic: jsonBody.topic,
            type: parseThreadType(jsonBody.chatType),
            memberIds: memberIds,
        };

        return result;
    } else {
        throw new Error(
            `Get chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const addMemberToChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    memberTeamsIdToBeAdd: string,
    userAccessToken: string
): Promise<void> => {
    const url = getGraphApiChatMemberUrl(threadId);

    const body = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${memberTeamsIdToBeAdd}')`,
        visibleHistoryStartDateTime: "0001-01-01T00:00:00Z",
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode !== HttpStatusCode.CREATED) {
        throw new Error(
            `Add member to group chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const listMembersInChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    userAccessToken: string
): Promise<string[]> => {
    const url = getGraphApiChatMemberUrl(threadId);
    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("List members in chat thread failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const userList = jsonBody.value as any[];

        const result: string[] = [];

        for (const user of userList) {
            result.push(user.userId);
        }

        return result;
    } else {
        throw new Error(
            `List members in chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const removeMemberFromChatThreadAsync = async (
    http: IHttp,
    threadId: string,
    teamsUserId: string,
    userAccessToken: string
): Promise<void> => {
    const url = getGraphApiChatMemberRemoveUrl(threadId, teamsUserId);
    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.del(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        return;
    } else {
        throw new Error(
            `Remove member from chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const sendTextMessageToChatThreadAsync = async (
    http: IHttp,
    textMessage: string,
    threadId: string,
    userAccessToken: string
): Promise<SendMessageResponse> => {
    const url = getGraphApiMessageUrl(threadId);

    const body = {
        body: {
            content: textMessage,
            contentType: "html",
        },
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Send message to chat thread failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: SendMessageResponse = {
            messageId: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(
            `Send message to chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const sendFileMessageToChatThreadAsync = async (
    http: IHttp,
    textMessage: string,
    fileName: string,
    shareLink: string,
    threadId: string,
    userAccessToken: string
): Promise<SendMessageResponse> => {
    const url = getGraphApiMessageUrl(threadId);

    const body = {
        body: {
            content: `${textMessage} <a href=\"${shareLink}\" title=\"${shareLink}\" target=\"_blank\" rel=\"noreferrer noopener\">${fileName}</a>`,
            contentType: "html",
        },
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Send file message to chat thread failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: SendMessageResponse = {
            messageId: jsonBody.id,
        };

        return result;
    } else {
        throw new Error(
            `Send file message to chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const updateTextMessageInChatThreadAsync = async (
    http: IHttp,
    textMessage: string,
    messageId: string,
    threadId: string,
    userAccessToken: string
): Promise<void> => {
    const url = getGraphApiMessageUrl(threadId, messageId, true);

    const body = {
        body: {
            content: textMessage,
        },
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.patch(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        return;
    } else {
        throw new Error(
            `Update message in chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const deleteTextMessageInChatThreadAsync = async (
    http: IHttp,
    teamsUserId: string,
    messageId: string,
    threadId: string,
    userAccessToken: string
): Promise<void> => {
    const url = getGraphApiMessageDeleteUrl(teamsUserId, threadId, messageId);

    const httpRequest: IHttpRequest = {
        headers: {
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        return;
    } else {
        throw new Error(
            `Delete message in chat thread failed with http status code ${response.statusCode}.`
        );
    }
};

export const getMessageWithResourceStringAsync = async (
    http: IHttp,
    resourceString: string,
    userAccessToken: string
): Promise<GetMessageResponse> => {
    const url = getGraphApiResourceUrl(resourceString);

    const httpRequest: IHttpRequest = {
        headers: {
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Get message with resource string failed!");
        }

        const jsonBody = JSON.parse(responseBody);

        let attachments: Attachment[] | undefined = undefined;

        const jsonAttachments = jsonBody.attachments as any[];
        if (jsonAttachments && jsonAttachments.length > 0) {
            attachments = [];
            for (const jsonAttachment of jsonAttachments) {
                const attachment: Attachment = {
                    id: jsonAttachment.id,
                    contentType: jsonAttachment.contentType,
                    contentUrl: jsonAttachment.contentUrl,
                    name: jsonAttachment.name,
                };
                attachments.push(attachment);
            }
        }

        const messageType = parseMessageType(
            jsonBody.messageType,
            jsonBody.eventDetail
        );

        let memberIds: string[] | undefined = undefined;
        if (messageType === MessageType.SystemAddMembers) {
            memberIds = [];
            const jsonMembers = jsonBody.eventDetail.members as any[];
            for (const jsonMember of jsonMembers) {
                memberIds.push(jsonMember.id);
            }
        }

        const result: GetMessageResponse = {
            threadId: jsonBody.chatId,
            messageId: jsonBody.id,
            messageType: messageType,
            fromUserTeamsId: jsonBody.from?.user?.id,
            messageContentType: parseMessageContentType(
                jsonBody.body?.contentType
            ),
            messageContent: jsonBody.body?.content,
            attachments: attachments,
            memberIds: memberIds,
        };

        return result;
    } else {
        throw new Error(
            `Get message with resource string failed with http status code ${response.statusCode}.`
        );
    }
};

export const listSubscriptionsAsync = async (
    http: IHttp,
    userAccessToken: string
): Promise<string[] | undefined> => {
    const url = getGraphApiSubscriptionUrl();

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("List subscriptions failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        let result: string[] | undefined = undefined;

        const jsonValues = jsonBody.value as any[];
        if (jsonValues && jsonValues.length > 0) {
            result = [];
            for (const value of jsonValues) {
                result.push(value.id);
            }
        }

        return result;
    } else {
        throw new Error(
            `List subscriptions failed with http status code ${response.statusCode}.`
        );
    }
};

export const renewSubscriptionAsync = async (
    http: IHttp,
    subscriptionId: string,
    userAccessToken: string,
    expirationDateTime?: Date
): Promise<SubscriptionResponse> => {
    if (!expirationDateTime) {
        expirationDateTime = new Date();
        expirationDateTime.setSeconds(
            expirationDateTime.getSeconds() + SubscriptionMaxExpireTimeInSecond
        );
    }

    const url = getGraphApiSubscriptionOperationUrl(subscriptionId);

    const body = {
        expirationDateTime: expirationDateTime.toISOString(),
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
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Renew subscription failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: SubscriptionResponse = {
            subscriptionId: jsonBody.id,
            expirationTime: new Date(jsonBody.expirationDateTime),
        };

        return result;
    } else {
        throw new Error(
            `Renew subscription failed with http status code ${response.statusCode}.`
        );
    }
};

export const deleteSubscriptionAsync = async (
    http: IHttp,
    subscriptionId: string,
    userAccessToken: string
): Promise<void> => {
    const url = getGraphApiSubscriptionOperationUrl(subscriptionId);

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.del(url, httpRequest);

    if (response.statusCode === HttpStatusCode.NO_CONTENT) {
        return;
    } else {
        throw new Error(
            `Delete subscription failed with http status code ${response.statusCode}.`
        );
    }
};

export const subscribeToAllMessagesForOneUserAsync = async (
    http: IHttp,
    rocketChatUserId: string,
    teamsUserId: string,
    subscriberEndpointUrl: string,
    userAccessToken: string,
    expirationDateTime?: Date
): Promise<SubscriptionResponse> => {
    try {
        if (!expirationDateTime) {
            expirationDateTime = new Date();
            expirationDateTime.setSeconds(
                expirationDateTime.getSeconds() +
                    SubscriptionMaxExpireTimeInSecond
            );
        }

        const url = getGraphApiSubscriptionUrl();

        const body = {
            changeType: SupportedNotificationChangeTypes.join(","),
            notificationUrl: `${subscriberEndpointUrl}?userId=${rocketChatUserId}`,
            resource: `/users/${teamsUserId}/chats/getAllMessages`,
            includeResourceData: false,
            expirationDateTime: expirationDateTime.toISOString(),
        };

        const httpRequest: IHttpRequest = {
            headers: {
                "Content-Type": "application/json",
                Authorization: `Bearer ${userAccessToken}`,
            },
            content: JSON.stringify(body),
        };

        const response = await http.post(url, httpRequest);

        if (response.statusCode === HttpStatusCode.CREATED) {
            const responseBody = response.content;
            if (responseBody === undefined) {
                throw new Error("Subscribe to notification for user failed!");
            }

            const jsonBody = JSON.parse(responseBody);
            const result: SubscriptionResponse = {
                subscriptionId: jsonBody.id,
                expirationTime: new Date(jsonBody.expirationDateTime),
            };

            return result;
        } else {
            throw new Error(
                `Subscribe to notification for user failed with http status code ${response.statusCode}.`
            );
        }
    } catch (error) {
        return error;
    }
};

export const downloadOneDriveFileAsync = async (
    http: IHttp,
    encodedUrl: string,
    userAccessToken: string
): Promise<Buffer> => {
    const url = getGraphApiShareUrl(encodedUrl);

    const httpRequest: IHttpRequest = {
        headers: {
            Authorization: `Bearer ${userAccessToken}`,
        },
        encoding: null,
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const fileStr = response.content as string;
        const buff = Buffer.from(fileStr, "binary");
        return buff;
    } else {
        throw new Error(
            `Download one drive file failed with http status code ${response.statusCode}.`
        );
    }
};

export const uploadFileToOneDriveAsync = async (
    http: IHttp,
    fileName: string,
    fileMIMEType: string,
    fileSize: number,
    content: Buffer,
    userAccessToken: string
): Promise<UploadFileResponse | undefined> => {
    if (fileSize < 4096000) {
        // use direct upload
        const url = getGraphApiUploadToDriveUrl(fileName);

        const httpRequest: IHttpRequest = {
            headers: {
                Authorization: `Bearer ${userAccessToken}`,
                "Content-Type": fileMIMEType,
            },
            content: content as any as string,
        };

        const response = await http.put(url, httpRequest);

        if (response.statusCode === HttpStatusCode.CREATED) {
            const responseBody = response.content;
            if (responseBody === undefined) {
                throw new Error("Upload file to one drive failed!");
            }

            const jsonBody = JSON.parse(responseBody);
            const result: UploadFileResponse = {
                driveItemId: jsonBody.id,
                fileName: jsonBody.name,
                size: jsonBody.size,
            };

            return result;
        } else {
            throw new Error(
                `Upload file to one drive failed with http status code ${response.statusCode}.`
            );
        }
    } else {
        // TODO: implement resumable upload
        return undefined;
    }
};

export const shareOneDriveFileAsync = async (
    http: IHttp,
    oneDriveItemId: string,
    userAccessToken: string
): Promise<ShareOneDriveFileResponse> => {
    const url = getGraphApiShareOneDriveFileUrl(oneDriveItemId);

    const body = {
        type: "view",
        scope: "organization",
    };

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
        content: JSON.stringify(body),
    };

    const response = await http.post(url, httpRequest);

    if (response.statusCode === HttpStatusCode.CREATED) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Create share link for onedrive item failed!");
        }

        const jsonBody = JSON.parse(responseBody);
        const result: ShareOneDriveFileResponse = {
            shareId: jsonBody.id,
            shareLink: jsonBody.link.webUrl,
        };

        return result;
    } else {
        throw new Error(
            `Create share link for onedrive item failed with http status code ${response.statusCode}.`
        );
    }
};

export const getOneDriveFileLinkAsync = async (
    http: IHttp,
    oneDriveItemId: string,
    userAccessToken: string
): Promise<string> => {
    const url = getGraphApiOneDriveFileLinkUrl(oneDriveItemId);

    const httpRequest: IHttpRequest = {
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${userAccessToken}`,
        },
    };

    const response = await http.get(url, httpRequest);

    if (response.statusCode === HttpStatusCode.OK) {
        const responseBody = response.content;
        if (responseBody === undefined) {
            throw new Error("Get one drive file link failed!");
        }

        const jsonBody = JSON.parse(responseBody);

        const result: string = jsonBody.webUrl as string;

        return result;
    } else {
        throw new Error(
            `Get one drive file link failed with http status code ${response.statusCode}.`
        );
    }
};

const parseMessageType = (
    messageType: string,
    eventDetail?: any
): MessageType | undefined => {
    if (!messageType) {
        return undefined;
    }

    if (messageType === "message") {
        return MessageType.Message;
    } else {
        if (eventDetail) {
            if (
                eventDetail["@odata.type"] ===
                "#microsoft.graph.membersAddedEventMessageDetail"
            ) {
                return MessageType.SystemAddMembers;
            }
        }
    }

    return undefined;
};

const parseMessageContentType = (
    messageContentType: string
): MessageContentType | undefined => {
    if (!messageContentType) {
        return undefined;
    }

    if (messageContentType === "html") {
        return MessageContentType.Html;
    }

    return undefined;
};

const parseThreadType = (threadType: string): ThreadType | undefined => {
    if (!threadType) {
        return undefined;
    }

    if (threadType === "group") {
        return ThreadType.Group;
    }

    if (threadType === "oneOnOne") {
        return ThreadType.OneOnOne;
    }

    return undefined;
};
