import { IPersistence, IPersistenceRead, IRead } from "@rocket.chat/apps-engine/definition/accessors";
import { RocketChatAssociationModel, RocketChatAssociationRecord } from "@rocket.chat/apps-engine/definition/metadata";
import { TestEnvironment } from "./Const";

const MiscKeys = {
    ApplicationAccessToken: 'ApplicationAccessToken',
    UserAccessToken: 'UserAccessToken',
    DummyUser: 'DummyUser',
};

interface ApplicationAccessTokenModel {
    accessToken: string,
};

interface UserAccessTokenModel {
    userId: string,
    accessToken: string,
    refreshToken: string,
    expires: number,
    extExpires: number,
};

export interface DummyUserModel {
    rocketChatUserId: string,
    teamsUserId: string,
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
    userId: string,
    accessToken: string,
    refreshToken: string,
    expiresIn: number,
    extExpiresIn: number) : Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.UserAccessToken),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, userId),
    ];

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    const data : UserAccessTokenModel = {
        userId: userId,
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
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.DummyUser),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
    ];
    const data : DummyUserModel = {
        rocketChatUserId: rocketChatUserId,
        teamsUserId: teamsUserId,
    };

    await persis.updateByAssociations(associations, data, true);
};

export const checkDummyUserAsync = async (read: IRead, userId: string) : Promise<boolean> => {
    const data = await retrieveDummyUserAsync(read, userId);
    
    if (data === undefined || data === null) {
        return false;
    }

    return true;
};

export const retrieveDummyUserAsync = async (read: IRead, userId: string) : Promise<DummyUserModel | null> => {
    // Mock dummy user before find out how to create user with Rocket.Chat Apps Engine
    if (TestEnvironment.enable) {
        const mockDummyUser = TestEnvironment.mockDummyUsers.filter(user => user.rocketChatUserId === userId);
        if (mockDummyUser && mockDummyUser.length === 1) {
            return mockDummyUser[0];
        }
    }

    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.DummyUser),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, userId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one DummyUser record for user ${userId}`);
    }

    const data : DummyUserModel = results[0] as DummyUserModel;
    return data;
};

export const retrieveUserAccessTokenAsync = async (read: IRead, userId: string) : Promise<string | null> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.UserAccessToken),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, userId),
    ];

    const persistenceRead : IPersistenceRead = read.getPersistenceReader();
    const results = await persistenceRead.readByAssociations(associations);

    if (results === undefined || results === null || results.length == 0) {
        return null;
    }

    if (results.length > 1) {
        throw new Error(`More than one UserAccessToken record for user ${userId}`);
    }

    const data : UserAccessTokenModel = results[0] as UserAccessTokenModel;

    const now = new Date();
    const epochInSecond = Math.round(now.getTime() / 1000);

    console.log(data);

    if (epochInSecond > data.expires) {
        return null;
    }

    return data.accessToken;
};
