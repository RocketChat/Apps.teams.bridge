import { IPersistence, IPersistenceRead, IRead } from "@rocket.chat/apps-engine/definition/accessors";
import { RocketChatAssociationModel, RocketChatAssociationRecord } from "@rocket.chat/apps-engine/definition/metadata";
import { TestEnvironment } from "./Const";

const MiscKeys = {
    ApplicationAccessToken: 'ApplicationAccessToken',
    UserAccessToken: 'UserAccessToken',
    User: 'User',
    DummyUser: 'DummyUser',
    Subscription: 'Subscription',
    MessageIdMapping: 'MessageIdMapping',
};

interface ApplicationAccessTokenModel {
    accessToken: string,
};

interface UserAccessTokenModel {
    rocketChatUserId: string,
    accessToken: string,
    refreshToken: string,
    expires: number,
    extExpires: number,
};

export interface UserModel {
    rocketChatUserId: string,
    teamsUserId: string,
};

export interface SubscriptionModel {
    rocketChatUserId: string,
    subscriptionId: string,
    expires: number,
};

export interface MessageIdModel {
    rocketChatMessageId: string,
    teamsMessageId: string,
    teamsThreadId: string,
};

export const persistApplicationAccessTokenAsync = async (
    persis: IPersistence,
    accessToken: string) : Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.ApplicationAccessToken),
    ];
    const data : ApplicationAccessTokenModel = {
        accessToken: accessToken,
    };

    await persis.updateByAssociations(associations, data, true);
};

export const persistUserAccessTokenAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    accessToken: string,
    refreshToken: string,
    expiresIn: number,
    extExpiresIn: number) : Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.UserAccessToken),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    const data : UserAccessTokenModel = {
        rocketChatUserId: rocketChatUserId,
        accessToken: accessToken,
        refreshToken: refreshToken,
        expires: epochInSecond + expiresIn,
        extExpires: epochInSecond + extExpiresIn,
    };

    await persis.updateByAssociations(associations, data, true);
};

export const persistDummyUserAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    teamsUserId: string) : Promise<void> => {
    const associationsByRocketChatUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.DummyUser),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];
    const associationsByTeamsUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.DummyUser),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
    ];

    const data : UserModel = {
        rocketChatUserId: rocketChatUserId,
        teamsUserId: teamsUserId,
    };

    await persis.updateByAssociations(associationsByRocketChatUserId, data, true);
    await persis.updateByAssociations(associationsByTeamsUserId, data, true);
};

export const persistUserAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    teamsUserId: string) : Promise<void> => {
    const associationsByRocketChatUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.User),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];
    const associationsByTeamsUserId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.User),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
    ];
    const data : UserModel = {
        rocketChatUserId: rocketChatUserId,
        teamsUserId: teamsUserId,
    };

    await persis.updateByAssociations(associationsByRocketChatUserId, data, true);
    await persis.updateByAssociations(associationsByTeamsUserId, data, true);
};

export const persistSubscriptionAsync = async (
    persis: IPersistence,
    rocketChatUserId: string,
    subscriptionId: string,
    expirationTime: Date) : Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.Subscription),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];
    const data : SubscriptionModel = {
        rocketChatUserId: rocketChatUserId,
        subscriptionId: subscriptionId,
        expires: Math.round(expirationTime.getTime() / 1000),
    };

    await persis.updateByAssociations(associations, data, true);
};

export const persistMessageIdMappingAsync = async (
    persis: IPersistence,
    rocketChatMessageId: string,
    teamsMessageId: string,
    teamsThreadId: string) : Promise<void> => {
    const associationsByRocketChatMessageId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.MessageIdMapping),
        new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatMessageId),
    ];
    const associationsByTeamsMessageId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.MessageIdMapping),
        new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, teamsMessageId),
    ];
    const data : MessageIdModel = {
        rocketChatMessageId: rocketChatMessageId,
        teamsMessageId: teamsMessageId,
        teamsThreadId: teamsThreadId,
    };

    await persis.updateByAssociations(associationsByRocketChatMessageId, data, true);
    await persis.updateByAssociations(associationsByTeamsMessageId, data, true);
};

export const checkDummyUserByRocketChatUserIdAsync = async (read: IRead, rocketChatUserId: string) : Promise<boolean> => {
    const data = await retrieveDummyUserByRocketChatUserIdAsync(read, rocketChatUserId);
    
    if (data === undefined || data === null) {
        return false;
    }

    return true;
};

export const retrieveDummyUserByRocketChatUserIdAsync = async (read: IRead, rocketChatUserId: string) : Promise<UserModel | null> => {
    // Mock dummy user before find out how to create user with Rocket.Chat Apps Engine
    if (TestEnvironment.enable) {
        const mockDummyUser = TestEnvironment.mockDummyUsers.filter(user => user.rocketChatUserId === rocketChatUserId);
        if (mockDummyUser && mockDummyUser.length === 1) {
            return mockDummyUser[0];
        }
    }

    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.DummyUser),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one DummyUser record for Rocket.Chat user ${rocketChatUserId}`);
    }

    const data : UserModel = results[0] as UserModel;
    return data;
};

export const retrieveDummyUserByTeamsUserIdAsync = async (read: IRead, teamsUserId: string) : Promise<UserModel | null> => {
    // Mock dummy user before find out how to create user with Rocket.Chat Apps Engine
    if (TestEnvironment.enable) {
        const mockDummyUser = TestEnvironment.mockDummyUsers.filter(user => user.teamsUserId === teamsUserId);
        if (mockDummyUser && mockDummyUser.length === 1) {
            return mockDummyUser[0];
        }
    }

    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.DummyUser),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one DummyUser record for Teams user ${teamsUserId}`);
    }

    const data : UserModel = results[0] as UserModel;
    return data;
};

export const retrieveUserByRocketChatUserIdAsync = async (read: IRead, rocketChatUserId: string) : Promise<UserModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.User),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one User record for user ${rocketChatUserId}`);
    }

    const data : UserModel = results[0] as UserModel;
    return data;
};

export const retrieveUserByTeamsUserIdAsync = async (read: IRead, teamsUserId: string) : Promise<UserModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.User),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one User record for user ${teamsUserId}`);
    }

    const data : UserModel = results[0] as UserModel;
    return data;
};

export const retrieveUserAccessTokenAsync = async (read: IRead, rocketChatUserId: string) : Promise<string | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.UserAccessToken),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one UserAccessToken record for user ${rocketChatUserId}`);
    }

    const data : UserAccessTokenModel = results[0] as UserAccessTokenModel;

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    if (!data.expires || epochInSecond > data.expires) {
        return null;
    }

    return data.accessToken;
};

export const retrieveSubscriptionAsync = async (read: IRead, rocketChatUserId: string) : Promise<SubscriptionModel | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.Subscription),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one Subscription record for user ${rocketChatUserId}`);
    }

    const data : SubscriptionModel = results[0] as SubscriptionModel;

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    if (!data.expires || epochInSecond > data.expires) {
        return null;
    }

    return data;
};

export const retrieveMessageIdMappingByRocketChatMessageIdAsync = async (
    read: IRead,
    rocketChatMessageId: string) : Promise<MessageIdModel | null> => {
    const associationsByRocketChatMessageId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.MessageIdMapping),
        new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatMessageId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associationsByRocketChatMessageId);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one ID mapping record for message ${rocketChatMessageId}`);
    }

    const data : MessageIdModel = results[0] as MessageIdModel;

    return data;
};

export const retrieveMessageIdMappingByTeamsMessageIdAsync = async (
    read: IRead,
    teamsMessageId: string) : Promise<MessageIdModel | null> => {
    const associationsByTeamsMessageId: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.MessageIdMapping),
        new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, teamsMessageId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associationsByTeamsMessageId);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one ID mapping record for message ${teamsMessageId}`);
    }

    const data : MessageIdModel = results[0] as MessageIdModel;

    return data;
};
