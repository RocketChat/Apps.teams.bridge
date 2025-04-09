import {
    IPersistence,
    IPersistenceRead,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import { IMessage } from '@rocket.chat/apps-engine/definition/messages';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';
import { IRoom } from '@rocket.chat/apps-engine/definition/rooms';
import { IUser } from '@rocket.chat/apps-engine/definition/users';
import { randomBytes, createHmac } from 'crypto';
import * as sha256 from 'crypto-js/sha256';

const MiscKeys = {
    ApplicationAccessToken: 'ApplicationAccessToken',
    UserRegistration: 'UserAccessToken',
    User: 'User',
    DummyUser: 'DummyUser',
    Subscription: 'Subscription',
    MessageIdMapping: 'MessageIdMapping',
    Room: 'Room',
    TeamsUserProfile: 'TeamsUserProfile',
    OneDriveFile: 'OneDriveFile',
    LoginMessage: 'LoginMessage',
    BridgedMessage: 'BridgedMessage',
    BridgedMessageFootprint: 'BridgedMessageFootprint',
    WebhookSecret: 'webhook-secret',
};

interface ApplicationAccessTokenModel {
    accessToken: string;
}

interface UserRegistrationModel {
    rocketChatUserId: string;
    accessToken: string;
    refreshToken: string;
    expires: number;
    extExpires: number;
}

export interface UserModel {
    rocketChatUserId: string;
    teamsUserId: string;
}

export interface SubscriptionModel {
    rocketChatUserId: string;
    subscriptionId: string;
    expires: number;
}

export interface MessageIdModel {
    rocketChatMessageId: string;
    teamsMessageId: string;
    teamsThreadId: string;
}

export interface RoomModel {
    rocketChatRoomId: string;
    teamsThreadId?: string;
    bridgeUserRocketChatUserId?: string;
}

export interface TeamsUserProfileModel {
    displayName: string;
    givenName: string;
    surname: string;
    mail: string;
    teamsUserId: string;
}

export interface OneDriveFileModel {
    fileName: string;
    driveItemId: string;
}

export const persistApplicationAccessTokenAsync = async (
    persis: IPersistence,
    accessToken: string
): Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.ApplicationAccessToken
        ),
    ];
    const data: ApplicationAccessTokenModel = {
        accessToken,
    };

    await persis.updateByAssociations(associations, data, true);
};

export const persistUserAccessTokenAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    accessToken: string,
    refreshToken: string,
    expiresIn: number,
    extExpiresIn: number
): Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.UserRegistration
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    const data: UserRegistrationModel = {
        rocketChatUserId,
        accessToken,
        refreshToken,
        expires: epochInSecond + expiresIn,
        extExpires: epochInSecond + extExpiresIn,
    };

    await persis.updateByAssociations(associations, data, true);
};

export const persistDummyUserAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    teamsUserId: string
): Promise<void> => {
    const associationsByRocketChatUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.DummyUser
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];
    const associationsByTeamsUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.DummyUser
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            teamsUserId
        ),
    ];

    const data: UserModel = {
        rocketChatUserId,
        teamsUserId,
    };

    await persis.updateByAssociations(
        associationsByRocketChatUserId,
        data,
        true
    );
    await persis.updateByAssociations(associationsByTeamsUserId, data, true);
};

export const persistUserAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    teamsUserId: string
): Promise<void> => {
    const associationsByRocketChatUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.User
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];
    const associationsByTeamsUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.User
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            teamsUserId
        ),
    ];
    const data: UserModel = {
        rocketChatUserId,
        teamsUserId,
    };

    await persis.updateByAssociations(
        associationsByRocketChatUserId,
        data,
        true
    );
    await persis.updateByAssociations(associationsByTeamsUserId, data, true);
};

export const persistSubscriptionAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    subscriptionId: string,
    expirationTime: Date
): Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.Subscription
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];
    const data: SubscriptionModel = {
        rocketChatUserId,
        subscriptionId,
        expires: Math.round(expirationTime.getTime() / 1000),
    };

    await persis.updateByAssociations(associations, data, true);
};

export const persistMessageIdMappingAsync = async (
    persis: IPersistence,
    rocketChatMessageId: string,
    teamsMessageId: string,
    teamsThreadId: string
): Promise<void> => {
    const associationsByRocketChatMessageId: Array<RocketChatAssociationRecord> =
        [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                MiscKeys.MessageIdMapping
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MESSAGE,
                rocketChatMessageId
            ),
        ];
    const associationsByTeamsMessageId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.MessageIdMapping
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MESSAGE,
            teamsMessageId
        ),
    ];
    const data: MessageIdModel = {
        rocketChatMessageId,
        teamsMessageId,
        teamsThreadId,
    };

    await persis.updateByAssociations(
        associationsByRocketChatMessageId,
        data,
        true
    );
    await persis.updateByAssociations(associationsByTeamsMessageId, data, true);
};

export const persistRoomAsync = async (
    persis: IPersistence,
    rocketChatRoomId: string,
    teamsThreadId?: string,
    bridgeUserRocketChatUserId?: string
): Promise<void> => {
    const associationsByRocketChatRoomId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.Room
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MESSAGE,
            rocketChatRoomId
        ),
    ];

    const data: RoomModel = {
        rocketChatRoomId,
        teamsThreadId,
        bridgeUserRocketChatUserId,
    };

    await persis.updateByAssociations(
        associationsByRocketChatRoomId,
        data,
        true
    );

    if (teamsThreadId) {
        const associationsByTeamsThreadId: Array<RocketChatAssociationRecord> =
            [
                new RocketChatAssociationRecord(
                    RocketChatAssociationModel.MISC,
                    MiscKeys.Room
                ),
                new RocketChatAssociationRecord(
                    RocketChatAssociationModel.MESSAGE,
                    teamsThreadId
                ),
            ];

        await persis.updateByAssociations(
            associationsByTeamsThreadId,
            data,
            true
        );
    }
};

export const persistTeamsUserProfileAsync = async (
    persis: IPersistence,
    displayName: string,
    givenName: string,
    surname: string,
    mail: string,
    teamsUserId: string
): Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.TeamsUserProfile
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            teamsUserId
        ),
    ];

    const data: TeamsUserProfileModel = {
        displayName,
        givenName,
        surname,
        mail,
        teamsUserId,
    };

    await persis.updateByAssociations(associations, data, true);
};

export const persistOneDriveFileAsync = async (
    persis: IPersistence,
    fileName: string,
    driveItemId: string
): Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.OneDriveFile
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.FILE,
            fileName
        ),
    ];

    const data: OneDriveFileModel = {
        fileName,
        driveItemId,
    };

    await persis.updateByAssociations(associations, data, true);
};

export const checkDummyUserByRocketChatUserIdAsync = async (
    read: IRead,
    rocketChatUserId: string
): Promise<boolean> => {
    const data = await retrieveDummyUserByRocketChatUserIdAsync(
        read,
        rocketChatUserId
    );

    if (data === undefined || data === null) {
        return false;
    }

    return true;
};

export const retrieveDummyUserByRocketChatUserIdAsync = async (
    read: IRead,
    rocketChatUserId: string
): Promise<UserModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.DummyUser
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one DummyUser record for Rocket.Chat user ${rocketChatUserId}`
        );
    }

    const data: UserModel = results[0] as UserModel;
    return data;
};

export const retrieveDummyUserByTeamsUserIdAsync = async (
    read: IRead,
    teamsUserId: string
): Promise<UserModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.DummyUser
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            teamsUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one DummyUser record for Teams user ${teamsUserId}`
        );
    }

    const data: UserModel = results[0] as UserModel;
    return data;
};

export const retrieveUserByRocketChatUserIdAsync = async (
    read: IRead,
    rocketChatUserId: string
): Promise<UserModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.User
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one User record for user ${rocketChatUserId}`
        );
    }

    const data: UserModel = results[0] as UserModel;
    return data;
};

export const retrieveUserByTeamsUserIdAsync = async (
    read: IRead,
    teamsUserId: string
): Promise<UserModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.User
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            teamsUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one User record for user ${teamsUserId}`);
    }

    const data: UserModel = results[0] as UserModel;
    return data;
};

export const retrieveUserAccessTokenAsync = async (
    read: IRead,
    persistence: IPersistence,
    rocketChatUserId: string
): Promise<string | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.UserRegistration
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one UserAccessToken record for user ${rocketChatUserId}`
        );
    }

    const data: UserRegistrationModel = results[0] as UserRegistrationModel;

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    if (!data.expires || epochInSecond > data.expires) {
        await saveLoginMessageSentStatus({
            persistence,
            rocketChatUserId,
            wasSent: false,
        });
        return null;
    }

    return data.accessToken;
};

export const retrieveAllUsersAccessTokenAsync = async (
    read: IRead,
): Promise<Array<string> | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.UserRegistration
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations) as Array<UserRegistrationModel>;

    if (results == null || results.length == 0) {
        return null;
    }

    return results.map((result) => result.accessToken);
};

export const retrieveUserRefreshTokenAsync = async (
    read: IRead,
    persistence: IPersistence,
    rocketChatUserId: string
): Promise<string | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.UserRegistration
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one UserAccessToken record for user ${rocketChatUserId}`
        );
    }

    const data: UserRegistrationModel = results[0] as UserRegistrationModel;

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    if (!data.extExpires || epochInSecond > data.extExpires) {
        await saveLoginMessageSentStatus({
            persistence,
            rocketChatUserId,
            wasSent: false,
        });

        return null;
    }

    return data.refreshToken;
};

export const retrieveAllUserRegistrationsAsync = async (
    read: IRead
): Promise<Array<UserRegistrationModel> | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.UserRegistration
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    return (results as Array<UserRegistrationModel>).filter((registration) => {
        return registration.extExpires && epochInSecond <= registration.extExpires;
    });

};

export const retrieveSubscriptionAsync = async (
    read: IRead,
    rocketChatUserId: string
): Promise<SubscriptionModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.Subscription
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one Subscription record for user ${rocketChatUserId}`
        );
    }

    const data: SubscriptionModel = results[0] as SubscriptionModel;

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    if (!data.expires || epochInSecond > data.expires) {
        return null;
    }

    return data;
};

export const retrieveMessageIdMappingByRocketChatMessageIdAsync = async (
    read: IRead,
    rocketChatMessageId: string
): Promise<MessageIdModel | null> => {
    const associationsByRocketChatMessageId: Array<RocketChatAssociationRecord> =
        [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                MiscKeys.MessageIdMapping
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MESSAGE,
                rocketChatMessageId
            ),
        ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(
        associationsByRocketChatMessageId
    );

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one ID mapping record for message ${rocketChatMessageId}`
        );
    }

    const data: MessageIdModel = results[0] as MessageIdModel;

    return data;
};

export const retrieveMessageIdMappingByTeamsMessageIdAsync = async (
    read: IRead,
    teamsMessageId: string
): Promise<MessageIdModel | null> => {
    const associationsByTeamsMessageId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.MessageIdMapping
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MESSAGE,
            teamsMessageId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(
        associationsByTeamsMessageId
    );

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one ID mapping record for message ${teamsMessageId}`
        );
    }

    const data: MessageIdModel = results[0] as MessageIdModel;

    return data;
};

export const retrieveRoomByRocketChatRoomIdAsync = async (
    read: IRead,
    rocketChatRoomId: string
): Promise<RoomModel | null> => {
    const associationsByRocketChatRoomId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.Room
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MESSAGE,
            rocketChatRoomId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(
        associationsByRocketChatRoomId
    );

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one Room record for room ${rocketChatRoomId}`
        );
    }

    const data: RoomModel = results[0] as RoomModel;

    return data;
};

export const retrieveRoomByTeamsThreadIdAsync = async (
    read: IRead,
    teamsThreadId: string
): Promise<RoomModel | null> => {
    const associationsByRocketChatRoomId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.Room
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MESSAGE,
            teamsThreadId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(
        associationsByRocketChatRoomId
    );

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one Room record for Teams thread ${teamsThreadId}`
        );
    }

    const data: RoomModel = results[0] as RoomModel;

    return data;
};

export const retrieveAllTeamsUserProfilesAsync = async (
    read: IRead
): Promise<Array<TeamsUserProfileModel> | null> => {
    const association = new RocketChatAssociationRecord(
        RocketChatAssociationModel.MISC,
        MiscKeys.TeamsUserProfile
    );

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociation(association);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    const data: Array<TeamsUserProfileModel> =
        results as Array<TeamsUserProfileModel>;

    return data;
};

export const retrieveOneDriveFileAsync = async (
    read: IRead,
    fileName: string
): Promise<OneDriveFileModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.OneDriveFile
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.FILE,
            fileName
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(
            `More than one OneDrive file record for file ${fileName}`
        );
    }

    const data: OneDriveFileModel = results[0] as OneDriveFileModel;

    return data;
};

export const deleteUserAccessTokenAsync = async (
    persis: IPersistence,
    rocketChatUserId: string
): Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.UserRegistration
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    await persis.removeByAssociations(associations);
};

export const deleteUserAsync = async (
    read: IRead,
    persis: IPersistence,
    rocketChatUserId: string
): Promise<void> => {
    const user = await retrieveUserByRocketChatUserIdAsync(
        read,
        rocketChatUserId
    );
    if (!user) {
        return;
    }

    const associationsByRocketChatUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.User
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            user.rocketChatUserId
        ),
    ];
    const associationsByTeamsUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.User
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            user.teamsUserId
        ),
    ];

    await persis.removeByAssociations(associationsByRocketChatUserId);
    await persis.removeByAssociations(associationsByTeamsUserId);
};

export const debugCleanAllRoomAsync = async (persis: IPersistence) => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.Room
        ),
    ];

    await persis.removeByAssociations(associations);
};

export type LoginMessageStatus = {
    isLoginMessageSent: boolean;
    rocketChatUserId: string;
};

export const saveLoginMessageSentStatus = async ({
    persistence,
    rocketChatUserId,
    wasSent,
}: {
    persistence: IPersistence;
    rocketChatUserId: string;
    wasSent: boolean;
}) => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.LoginMessage
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const data: LoginMessageStatus = {
        isLoginMessageSent: wasSent,
        rocketChatUserId,
    };

    await persistence.updateByAssociations(associations, data, true);
};

export const retrieveLoginMessageSentStatus = async ({
    read,
    rocketChatUserId,
}: {
    read: IRead;
    rocketChatUserId: string;
}): Promise<boolean> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.LoginMessage
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();

    const result = (await persistenceRead.readByAssociations(
        associations
    )) as unknown as Array<LoginMessageStatus>;

    if (!result) {
        return false;
    }

    return !!result[0]?.isLoginMessageSent;
};

type LastBridgedMessageInfo = {
    senderId: string,
    text: string,
    roomId: string,
    fileName: string,
}

export const getLastBridgedMessage = async ({
    read,
    rocketChatUserId,
}: {
    read: IRead;
    rocketChatUserId: string;
}): Promise<LastBridgedMessageInfo | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.BridgedMessage
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    const persistenceRead: IPersistenceRead = read.getPersistenceReader();

    const result = (await persistenceRead.readByAssociations(
        associations
    )) as unknown as Array<LastBridgedMessageInfo>;

    if (!result) {
        return null;
    }

    return result[0];
};

const simpleHash = (input: string): string => sha256(input).toString()

const calculatePriorityDate = (message: IMessage) => {
    const currentDate = new Date();
    let priorityDate = currentDate;

    if (message.updatedAt) {
      const updatedAtDate = new Date(message.updatedAt);
      if (!isNaN(updatedAtDate.getTime())) {
        priorityDate = updatedAtDate;
      }
    } else if (message.createdAt) {
      const createdAtDate = new Date(message.createdAt);
      if (!isNaN(createdAtDate.getTime())) {
        priorityDate = createdAtDate;
      }
    }

    return priorityDate;
  }

export const generateMessageFootprint = (message: IMessage, room: IRoom, sender: IUser): string => {
    const dateObject = calculatePriorityDate(message);
    dateObject.setSeconds(0);
    dateObject.setMilliseconds(0);

    const timestamp = dateObject.getTime();
    const messageProperties = `${message.text}${timestamp}`;
    const roomProperties = `${room.id}${room.type}${room.displayName}`;

    const senderProperties = `${sender.id}${sender.username}`;

    const combinedProperties = `${messageProperties}${roomProperties}${senderProperties}`;

    return simpleHash(combinedProperties);
}

export const saveLastBridgedMessageFootprint = async ({
    persistence,
    rocketChatUserId,
    messageFootprint,
}: {
    messageFootprint: string
    persistence: IPersistence,
    rocketChatUserId: string;
}): Promise<string> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.BridgedMessageFootprint
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    return persistence.updateByAssociations(
        associations,
        { messageFootprint, createdAt: new Date().toString() },
        true
    );
};

export type MessageFootprintInfo = {
    messageFootprint: string
    createdAt: string
}

export const getLastBridgedMessageFootprint = async (
    {
        read,
        rocketChatUserId,
    }: {
        read: IRead,
        rocketChatUserId: string;
    }
): Promise<MessageFootprintInfo> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.BridgedMessageFootprint
        ),
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.USER,
            rocketChatUserId
        ),
    ];

    return (await read.getPersistenceReader().readByAssociations(associations)).shift() as MessageFootprintInfo
}

export const doesMessageFootPrintExists = async (
    {
        currentMessageFootprint,
        storedMessageFootprintInfo
    }:
    {
        currentMessageFootprint: string,
        storedMessageFootprintInfo: MessageFootprintInfo
    }
) => {
    return currentMessageFootprint === storedMessageFootprintInfo.messageFootprint
};

export const getMessageFootPrintExistenceInfo = async (message: IMessage, read: IRead):
    Promise<{
        itDoesMessageFootprintExists: boolean,
        messageFootprint: string
    } | undefined> => {

    try {
        const rocketChatUserId =  message.sender.id
        const storedMessageFootprintInfo = await getLastBridgedMessageFootprint({
            rocketChatUserId,
            read,
        });

        if (!storedMessageFootprintInfo) return;

        const currentMessageFootprint = generateMessageFootprint(message, message.room, message.sender);

        // Check if the message's footprint exists in the database and returns
        const itDoesMessageFootprintExists = await doesMessageFootPrintExists({
            currentMessageFootprint, storedMessageFootprintInfo
        });

        return {
            itDoesMessageFootprintExists,
            messageFootprint: currentMessageFootprint
        }

    } catch (error) {
        console.error("An error occured when trying to get message footprint info", error)
        return
    }

}

export const createWebhookSecret = async ({ persistence }: { persistence: IPersistence }) => {
    const associations = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.WebhookSecret
        ),
    ];

    const secretLength = 16;
    const secret = randomBytes(secretLength).toString("hex");
    await persistence.createWithAssociations({ secret }, associations);
    return secret;
}

export const getWebhookSecret = async ({
    persistenceRead
}: {
    persistenceRead: IPersistenceRead;
}): Promise<string | null> => {
    const associations = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.WebhookSecret
        ),
    ];
    const [record] = await persistenceRead.readByAssociations(associations);
    return (record as any)?.secret || null;
}

export const getOrCreatWebhookSecret = async ({
    persistenceRead,
    persistenceWrite,
}: {
    persistenceRead: IPersistenceRead;
    persistenceWrite: IPersistence;
}): Promise<string> => {
    const associations = [
        new RocketChatAssociationRecord(
            RocketChatAssociationModel.MISC,
            MiscKeys.WebhookSecret
        ),
    ];
    const [record] = await persistenceRead.readByAssociations(associations);
    const secret = (record as any)?.secret || null;

    if (secret) {
        return secret;
    }
    return await createWebhookSecret({ persistence: persistenceWrite });
};

export const getSubscriptionStateHashForUser = async (
    persistenceRead: IPersistenceRead,
    persistenceWrite: IPersistence,
    user: Pick<UserModel, "rocketChatUserId">
) => {
    const webhookSecret = await getOrCreatWebhookSecret({
        persistenceRead,
        persistenceWrite,
    });

    const hmac = createHmac("sha256", webhookSecret).update(
        user.rocketChatUserId
    );

    return hmac.digest("hex");
};
