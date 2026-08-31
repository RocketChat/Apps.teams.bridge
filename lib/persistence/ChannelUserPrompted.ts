import type { IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { RocketChatAssociationModel, RocketChatAssociationRecord } from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'ChannelUserPrompted';

// Remembers that the "new Teams user in channel — map them" prompt was already posted
// for one (channelId, teamsUserId) pair, so it fires once per new member.
export const ChannelUserPrompted = {
	assoc(channelId: string, teamsUserId: string): RocketChatAssociationRecord[] {
		return [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, `${channelId}:${teamsUserId}`),
		];
	},

	async set(persis: IPersistence, channelId: string, teamsUserId: string): Promise<void> {
		await persis.updateByAssociations(ChannelUserPrompted.assoc(channelId, teamsUserId), { prompted: true }, true);
	},

	async isSet(read: IRead, channelId: string, teamsUserId: string): Promise<boolean> {
		const results = await read.getPersistenceReader().readByAssociations(ChannelUserPrompted.assoc(channelId, teamsUserId));
		return !!results && results.length > 0;
	},
};
