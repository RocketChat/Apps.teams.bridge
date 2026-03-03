import {
    IPersistence,
    IPersistenceRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'TeamsBridgeSubscriptionRenewalJobState';

export interface SubscriptionRenewalJobState {
    lastStartedJobTimestamp: string | Date;
}

export const SubscriptionRenewalJob = {
    async persist(options: {
        persistence: IPersistence;
    } & SubscriptionRenewalJobState): Promise<SubscriptionRenewalJobState> {
        const { persistence, lastStartedJobTimestamp } = options;
        const associations = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
        ];
        const data: SubscriptionRenewalJobState = { lastStartedJobTimestamp };
        await persistence.updateByAssociations(associations, data, true);
        return data;
    },

    async find(options: {
        persistenceRead: IPersistenceRead;
    }): Promise<SubscriptionRenewalJobState | null> {
        const { persistenceRead } = options;
        const associations = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
        ];
        const [record] = await persistenceRead.readByAssociations(associations);
        return record ? (record as SubscriptionRenewalJobState) : null;
    },
};
