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
    SystemAddMembers = 'addMembers',
    SystemRemoveMembers = 'removeMembers',
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
    reactionType: string;
    displayName: string;
    reactionContentUrl: string | null;
    createdDateTime: string;
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

export type SubscriptionValue = {
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
