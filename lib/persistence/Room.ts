import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'Room';

export interface RoomModel {
    rocketChatRoomId: string;
    teamsThreadId?: string;
    isBridged?: boolean;
}

export const Room = {
    async persist(
        read: IRead,
        persis: IPersistence,
        rocketChatRoomId: string,
        teamsThreadId?: string,
    ): Promise<void> {
        const existing = await Room.findByRCRoomId(read, rocketChatRoomId);
        const byRoomId: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatRoomId),
        ];
        const data: RoomModel = { ...existing, rocketChatRoomId, teamsThreadId };
        await persis.updateByAssociations(byRoomId, data, true);

        if (teamsThreadId) {
            const byThreadId: Array<RocketChatAssociationRecord> = [
                new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
                new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, teamsThreadId),
            ];
            await persis.updateByAssociations(byThreadId, data, true);
        }
    },

    async findByRCRoomId(read: IRead, rocketChatRoomId: string): Promise<RoomModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatRoomId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one Room record for room ${rocketChatRoomId}`);
        }
        return results[0] as RoomModel;
    },

    async findByTeamsThreadId(read: IRead, teamsThreadId: string): Promise<RoomModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, teamsThreadId),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one Room record for Teams thread ${teamsThreadId}`);
        }
        return results[0] as RoomModel;
    },

    async isBridged(read: IRead, rocketChatRoomId: string): Promise<boolean> {
        const roomRecord = await Room.findByRCRoomId(read, rocketChatRoomId);
        return roomRecord?.isBridged === true;
    },

    async setBridgeActive(
        persistence: IPersistence,
        read: IRead,
        rocketChatRoomId: string,
        active: boolean,
    ): Promise<void> {
        const existing = await Room.findByRCRoomId(read, rocketChatRoomId);
        const data: RoomModel = {
            ...(existing ?? { rocketChatRoomId }),
            isBridged: active,
        };
        const byRoomId: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatRoomId),
        ];
        await persistence.updateByAssociations(byRoomId, data, true);

        if (data.teamsThreadId) {
            const byThreadId: Array<RocketChatAssociationRecord> = [
                new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
                new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, data.teamsThreadId),
            ];
            await persistence.updateByAssociations(byThreadId, data, true);
        }
    },

    async debugCleanAll(persis: IPersistence): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
        ];
        await persis.removeByAssociations(associations);
    },
};
