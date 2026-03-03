import {
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'OneDriveFile';

export interface OneDriveFileModel {
    fileName: string;
    driveItemId: string;
}

export const OneDriveFile = {
    async persist(persis: IPersistence, fileName: string, driveItemId: string): Promise<void> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.FILE, fileName),
        ];
        const data: OneDriveFileModel = { fileName, driveItemId };
        await persis.updateByAssociations(associations, data, true);
    },

    async find(read: IRead, fileName: string): Promise<OneDriveFileModel | null> {
        const associations: Array<RocketChatAssociationRecord> = [
            new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
            new RocketChatAssociationRecord(RocketChatAssociationModel.FILE, fileName),
        ];
        const results = await read.getPersistenceReader().readByAssociations(associations);
        if (!results || results.length === 0) {
            return null;
        }
        if (results.length > 1) {
            throw new Error(`More than one OneDrive file record for file ${fileName}`);
        }
        return results[0] as OneDriveFileModel;
    },
};
