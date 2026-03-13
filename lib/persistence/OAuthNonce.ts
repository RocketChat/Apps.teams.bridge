import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'OAuthNonce';

interface OAuthNonceRecord {
    rocketChatUserId: string;
    nonce: string;
    createdAt: number;
}

export const OAuthNonce = {
    async persist(
        persis: IPersistence,
        rocketChatUserId: string,
        nonce: string,
    ): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        await persis.updateByAssociations(associations, { rocketChatUserId, nonce, createdAt: Date.now() }, true);
    },

    async findAndDelete(
        read: IRead,
        persis: IPersistence,
        rocketChatUserId: string,
    ): Promise<string | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        await persis.removeByAssociations(associations);
        return (results[0] as OAuthNonceRecord).nonce ?? null;
    },

    async deleteStale(
        read: IRead,
        persis: IPersistence,
    ): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations) as OAuthNonceRecord[];
        if (!results || results.length === 0) {
            return;
        }
        const cutoff = Date.now() - 10 * 60 * 1000;
        await Promise.all(
            results
                .filter((r) => r.createdAt < cutoff)
                .map((r) => persis.removeByAssociations([
                    new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
                    new RocketChatAssociationRecord(RocketChatAssociationModel.USER, r.rocketChatUserId),
                ])),
        );
    },
};
