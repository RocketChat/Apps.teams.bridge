import {
    IPersistence,
    IPersistenceRead,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

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
        return null;
    }

    return data.accessToken;
};

export const retrieveUserRefreshTokenAsync = async (
    read: IRead,
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

    const data: Array<UserRegistrationModel> = [];

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);
    for (const result of results) {
        const registration = result as UserRegistrationModel;
        if (
            !registration.extExpires ||
            epochInSecond > registration.extExpires
        ) {
            continue;
        }
        data.push(registration);
    }

    return data;
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

    await persistence.updateByAssociations(
        associations,
        {
            isLoginMessageSent: true,
            rocketChatUserId,
        },
        wasSent
    );
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
    )) as unknown as Array<boolean>;

    if (!result) {
        return false;
    }

    return result[0];
};
