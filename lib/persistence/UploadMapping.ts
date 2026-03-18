import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'AttachmentMapping';

export interface UploadMappingModel {
    rocketchatUploadId: string;
    teamsMessageId: string;
    teamsThreadId: string;
    teamsAttachmentId: string;
    relayedByAppUser?: boolean;
}

export const UploadMapping = {
    async persist(options: {
        persistence: IPersistence;
        rocketchatUploadId: string;
        teamsMessageId: string;
        teamsThreadId: string;
        teamsAttachmentId: string;
        relayedByAppUser?: boolean;
    }): Promise<void> {
        const { persistence, rocketchatUploadId, teamsMessageId, teamsThreadId, teamsAttachmentId, relayedByAppUser } = options;
        const byUploadId: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, rocketchatUploadId),
        ];
        const byTeamsMessageId: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, teamsMessageId),
        ];
        const data: UploadMappingModel = {
            rocketchatUploadId,
            teamsMessageId,
            teamsThreadId,
            teamsAttachmentId,
            ...(relayedByAppUser !== undefined ? { relayedByAppUser } : {}),
        };
        await persistence.updateByAssociations(byUploadId, data, true);
        await persistence.updateByAssociations(byTeamsMessageId, data, true);
    },

    async findOneByRCUploadId(read: IRead, rocketchatUploadId: string): Promise<Array<UploadMappingModel>> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, rocketchatUploadId),
        ];
        const data = await read.getPersistenceReader().readByAssociations(associations) as Array<UploadMappingModel>;
        return data && data.length > 0 ? data : [];
    },

    async findAllByRCUploadId(read: IRead, rocketchatUploadId: string): Promise<Array<UploadMappingModel>> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, rocketchatUploadId),
        ];
        const data = (await read.getPersistenceReader().readByAssociations(associations))?.[0] as UploadMappingModel | undefined;
        if (data?.teamsMessageId) {
            return UploadMapping.findByTeamsMessageId(read, data.teamsMessageId);
        }
        return [];
    },

    async findByTeamsMessageId(read: IRead, teamsMessageId: string): Promise<Array<UploadMappingModel>> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, teamsMessageId),
        ];
        const data = await read.getPersistenceReader().readByAssociations(associations) as Array<UploadMappingModel>;
        return data && data.length > 0 ? data : [];
    },

    async delete(options: {
        persistence: IPersistence;
        rocketchatUploadId: string;
        teamsMessageId: string;
    }): Promise<void> {
        const { persistence, rocketchatUploadId, teamsMessageId } = options;
        const byUploadId: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, rocketchatUploadId),
        ];
        const byTeamsMessageId: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, teamsMessageId),
        ];
        await Promise.all([
            persistence.removeByAssociations(byUploadId),
            persistence.removeByAssociations(byTeamsMessageId),
        ]);
    },
};
