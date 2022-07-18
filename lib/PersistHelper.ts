import { IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { RocketChatAssociationModel, RocketChatAssociationRecord } from "@rocket.chat/apps-engine/definition/metadata";

const MiscKeys = {
    ApplicationAccessToken: 'ApplicationAccessToken',
    UserAccessToken: 'UserAccessToken',
};

interface ApplicationAccessTokenModel {
    accessToken: string,
};

interface UserAccessTokenModel {
    userId: string,
    accessToken: string,
    refreshToken: string,
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
    refreshToken: string) : Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.UserAccessToken),
        new RocketChatAssociationRecord(RocketChatAssociationModel.USER, userId),
    ];
    const data : UserAccessTokenModel = {
        userId: userId,
        accessToken: accessToken,
        refreshToken: refreshToken,
    };

    await persis.updateByAssociations(associations, data, true);
};
