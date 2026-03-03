import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'Subscription';

export interface SubscriptionModel {
    rocketChatUserId: string;
    subscriptionId: string;
    expires: number;
}

export const Subscription = {
    async persist(
        persis: IPersistence,
        rocketChatUserId: string,
        subscriptionId: string,
        expirationTime: Date,
    ): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        const data: SubscriptionModel = {
            rocketChatUserId,
            subscriptionId,
            expires: Math.round(expirationTime.getTime() / 1000),
        };
        await persis.updateByAssociations(associations, data, true);
    },

    async findByRCUserId(read: IRead, rocketChatUserId: string): Promise<SubscriptionModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one Subscription record for user ${rocketChatUserId}`);
        }
        const data = results[0] as SubscriptionModel;
        const epochInSecond = Math.round(new Date().getTime() / 1000);
        if (!data.expires || epochInSecond > data.expires) {
            return null;
        }
        return data;
    },
};
