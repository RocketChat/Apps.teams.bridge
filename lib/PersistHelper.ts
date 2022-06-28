import { IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { RocketChatAssociationModel, RocketChatAssociationRecord } from "@rocket.chat/apps-engine/definition/metadata";

const MiscKeys = {
    ApplicationAccessToken: 'ApplicationAccessToken',
};

interface ApplicationAccessTokenModel {
    token: string,
};

export const persistApplicationAccessTokenAsync = async (
    persis: IPersistence,
    applicationAccessToken: string) : Promise<void> => {
    const associations: Array<RocketChatAssociationRecord> = [
        new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, MiscKeys.ApplicationAccessToken),
    ];
    const data : ApplicationAccessTokenModel = {
        token: applicationAccessToken,
    };

    await persis.updateByAssociations(associations, data, true);
};
