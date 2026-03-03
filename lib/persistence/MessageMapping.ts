import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'MessageIdMapping';

export interface MessageMappingModel {
    rocketChatMessageId: string;
    teamsMessageId: string;
    teamsThreadId: string;
}

export const MessageMapping = {
    async persist(options: {
        persistence: IPersistence;
        rocketChatMessageId: string;
        teamsMessageId: string;
        teamsThreadId: string;
    }): Promise<void> {
        const { persistence, rocketChatMessageId, teamsMessageId, teamsThreadId } = options;
        const byRC: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatMessageId),
        ];
        const byTeams: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, teamsMessageId),
        ];
        const data: MessageMappingModel = { rocketChatMessageId, teamsMessageId, teamsThreadId };
        await persistence.updateByAssociations(byRC, data, true);
        await persistence.updateByAssociations(byTeams, data, true);
    },

    async findByRCMessageId(read: IRead, rocketChatMessageId: string): Promise<MessageMappingModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatMessageId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one ID mapping record for message ${rocketChatMessageId}`);
        }
        return results[0] as MessageMappingModel;
    },

    async findByTeamsMessageId(read: IRead, teamsMessageId: string): Promise<MessageMappingModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, teamsMessageId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one ID mapping record for message ${teamsMessageId}`);
        }
        return results[0] as MessageMappingModel;
    },

    async delete(options: {
        persistence: IPersistence;
        rocketChatMessageId: string;
        teamsMessageId: string;
    }): Promise<void> {
        const { persistence, rocketChatMessageId, teamsMessageId } = options;
        const byRC: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatMessageId),
        ];
        const byTeams: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, teamsMessageId),
        ];
        await Promise.all([
            persistence.removeByAssociations(byRC),
            persistence.removeByAssociations(byTeams),
        ]);
    },
};
