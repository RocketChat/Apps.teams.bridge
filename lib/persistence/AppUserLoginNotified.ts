import {
    IPersistence,
    IPersistenceRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'AppUserLoginNotified';

export const AppUserLoginNotified = {
    async set(persistence: IPersistence, roomId: string): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.ROOM,
                roomId,
            ),
        ];
        await persistence.updateByAssociations(
            associations,
            { notified: true, date: new Date() },
            true,
        );
    },

    async isSet(
        persistenceRead: IPersistenceRead,
        roomId: string,
    ): Promise<boolean> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.ROOM,
                roomId,
            ),
        ];
        const results = await persistenceRead.readByAssociations(associations);
        return results.length > 0;
    },

    async isSetToday(
        persistenceRead: IPersistenceRead,
        roomId: string,
    ): Promise<boolean> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.ROOM,
                roomId,
            ),
        ];
        const results = await persistenceRead.readByAssociations(associations) as Array<{ date: string; notified: boolean }>;

        const today = new Date();
        const isToday = (date: Date) =>
            date.getDate() === today.getDate() &&
            date.getMonth() === today.getMonth() &&
            date.getFullYear() === today.getFullYear();

        try {
            return results.some((result) => {
                const data = new Date(result.date);
                return result.notified && isToday(data);
            });
        } catch (error) {
            console.error('Error parsing date from persistence:', error);
            return false;
        }
    },

    /** Clears notification flags for all rooms — called when the app user logs in. */
    async clearAll(persistence: IPersistence): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
        ];
        await persistence.removeByAssociations(associations);
    },
};
