import { IPersistence } from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'ApplicationAccessToken';

interface AppTokenModel {
    accessToken: string;
}

export const AppToken = {
    async persist(persis: IPersistence, accessToken: string): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
        ];
        const data: AppTokenModel = { accessToken };
        await persis.updateByAssociations(associations, data, true);
    },
};
