import type { IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { RocketChatAssociationModel, RocketChatAssociationRecord } from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'ApplicationAccessToken';

interface AppTokenModel {
	accessToken: string;
	expires: number; // epoch seconds
}

export const AppToken = {
	async persist(persis: IPersistence, accessToken: string, expires: number): Promise<void> {
		const associations: Array<RocketChatAssociationRecord> = [new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY)];
		const data: AppTokenModel = { accessToken, expires };
		await persis.updateByAssociations(associations, data, true);
	},

	async find(read: IRead): Promise<AppTokenModel | null> {
		const associations: Array<RocketChatAssociationRecord> = [new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY)];
		const results = await read.getPersistenceReader().readByAssociations(associations);
		if (!results || results.length === 0) {
			return null;
		}
		return results[0] as AppTokenModel;
	},
};
