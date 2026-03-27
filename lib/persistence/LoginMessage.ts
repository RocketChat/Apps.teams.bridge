import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'LoginMessage';

export type LoginMessageStatus = {
    isLoginMessageSent: boolean;
    rocketChatUserId: string;
};

export const LoginMessage = {
    async save(options: {
        persistence: IPersistence;
        rocketChatUserId: string;
        wasSent: boolean;
    }): Promise<void> {
        const { persistence, rocketChatUserId, wasSent } = options;
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        const data: LoginMessageStatus = { isLoginMessageSent: wasSent, rocketChatUserId };
        await persistence.updateByAssociations(associations, data, true);
    },

    async get(options: {
        read: IRead;
        rocketChatUserId: string;
    }): Promise<boolean> {
        const { read, rocketChatUserId } = options;
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        const result = (await read.getPersistenceReader().readByAssociations(associations)) as unknown as Array<LoginMessageStatus>;
        if (!result) {
            return false;
        }
        return !!result[0]?.isLoginMessageSent;
    },

    async delete(
        persistence: IPersistence,
        rocketChatUserId: string,
    ): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        await persistence.removeByAssociations(associations);
    },
};
