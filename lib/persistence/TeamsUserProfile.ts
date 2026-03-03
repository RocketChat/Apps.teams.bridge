import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'TeamsUserProfile';

export interface TeamsUserProfileModel {
    displayName: string;
    givenName: string;
    surname: string;
    mail: string;
    teamsUserId: string;
}

export const TeamsUserProfile = {
    async persist(
        persis: IPersistence,
        displayName: string,
        givenName: string,
        surname: string,
        mail: string,
        teamsUserId: string,
    ): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
        ];
        const data: TeamsUserProfileModel = { displayName, givenName, surname, mail, teamsUserId };
        await persis.updateByAssociations(associations, data, true);
    },

    async findAll(read: IRead): Promise<Array<TeamsUserProfileModel> | null> {
        const association = new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY);
        const results = await read.getPersistenceReader().readByAssociation(association);
        if (!results || results.length === 0) {
            return null;
        }
        return results as Array<TeamsUserProfileModel>;
    },

    async findByTeamsUserId(read: IRead, teamsUserId: string): Promise<TeamsUserProfileModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        return results[0] as TeamsUserProfileModel;
    },
};
