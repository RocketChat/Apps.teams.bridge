import { createHmac, randomBytes } from 'crypto';

import type { IPersistence, IPersistenceRead } from '@rocket.chat/apps-engine/definition/accessors';
import { RocketChatAssociationModel, RocketChatAssociationRecord } from '@rocket.chat/apps-engine/definition/metadata';

import type { UserModel } from './UserMapping';

const KEY = 'webhook-secret';

export const WebhookSecret = {
	async create(options: { persistence: IPersistence }): Promise<string> {
		const { persistence } = options;
		const associations = [new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY)];
		const secret = randomBytes(16).toString('hex');
		await persistence.createWithAssociations({ secret }, associations);
		return secret;
	},

	async get(options: { persistenceRead: IPersistenceRead }): Promise<string | null> {
		const { persistenceRead } = options;
		const associations = [new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY)];
		const [record] = await persistenceRead.readByAssociations(associations);
		return (record as any)?.secret || null;
	},

	async getOrCreate(options: { persistenceRead: IPersistenceRead; persistenceWrite: IPersistence }): Promise<string> {
		const existing = await WebhookSecret.get({ persistenceRead: options.persistenceRead });
		if (existing) {
			return existing;
		}
		return WebhookSecret.create({ persistence: options.persistenceWrite });
	},

	async getSubscriptionStateHash(
		persistenceRead: IPersistenceRead,
		persistenceWrite: IPersistence,
		user: Pick<UserModel, 'rocketChatUserId'>,
	): Promise<string> {
		const secret = await WebhookSecret.getOrCreate({ persistenceRead, persistenceWrite });
		return createHmac('sha256', secret).update(user.rocketChatUserId).digest('hex');
	},
};
