import type { IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { RocketChatAssociationModel, RocketChatAssociationRecord } from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'UserAccessToken';

export interface UserRegistrationModel {
	rocketChatUserId: string;
	accessToken: string;
	refreshToken: string;
	expires: number;
	extExpires: number;
}

export const UserRegistration = {
	async persist(
		persis: IPersistence,
		rocketChatUserId: string,
		accessToken: string,
		refreshToken: string,
		expiresIn: number,
		extExpiresIn: number,
	): Promise<void> {
		const associations: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
		];
		const epochInSecond = Math.round(new Date().getTime() / 1000);
		const data: UserRegistrationModel = {
			rocketChatUserId,
			accessToken,
			refreshToken,
			expires: epochInSecond + expiresIn,
			extExpires: epochInSecond + extExpiresIn,
		};
		await persis.updateByAssociations(associations, data, true);
	},

	async findByRCUserId(options: { read: IRead; rocketChatUserId: string }): Promise<UserRegistrationModel | null> {
		const { read, rocketChatUserId } = options;
		const associations: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
		];
		const results = await read.getPersistenceReader().readByAssociations(associations);
		if (!results || results.length === 0) {
			return null;
		}
		if (results.length > 1) {
			throw new Error(`More than one UserAccessToken record for user ${rocketChatUserId}`);
		}
		return results[0] as UserRegistrationModel;
	},

	async findAll(read: IRead): Promise<Array<UserRegistrationModel> | null> {
		const associations: Array<RocketChatAssociationRecord> = [new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY)];
		const results = await read.getPersistenceReader().readByAssociations(associations);
		if (!results || results.length === 0) {
			return null;
		}
		return results as Array<UserRegistrationModel>;
	},

	async delete(persis: IPersistence, rocketChatUserId: string): Promise<void> {
		const associations: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
		];
		await persis.removeByAssociations(associations);
	},
};
