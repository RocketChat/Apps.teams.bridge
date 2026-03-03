import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'User';

export interface UserModel {
    rocketChatUserId: string;
    teamsUserId: string;
}

export const UserMapping = {
    async persist(
        persis: IPersistence,
        rocketChatUserId: string,
        teamsUserId: string,
    ): Promise<void> {
        const byRC: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        const byTeams: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
        ];
        const data: UserModel = { rocketChatUserId, teamsUserId };
        await persis.updateByAssociations(byRC, data, true);
        await persis.updateByAssociations(byTeams, data, true);
    },

    async findByRCUserId(read: IRead, rocketChatUserId: string): Promise<UserModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one User record for user ${rocketChatUserId}`);
        }
        return results[0] as UserModel;
    },

    async findByTeamsUserId(read: IRead, teamsUserId: string): Promise<UserModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, teamsUserId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one User record for user ${teamsUserId}`);
        }
        return results[0] as UserModel;
    },

    async delete(read: IRead, persis: IPersistence, rocketChatUserId: string): Promise<void> {
        const user = await UserMapping.findByRCUserId(read, rocketChatUserId);
        if (!user) {
            return;
        }
        const byRC: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, user.rocketChatUserId),
        ];
        const byTeams: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.USER, user.teamsUserId),
        ];
        await persis.removeByAssociations(byRC);
        await persis.removeByAssociations(byTeams);
    },
};
