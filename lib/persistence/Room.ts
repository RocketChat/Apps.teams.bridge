import type { IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { RocketChatAssociationModel, RocketChatAssociationRecord } from '@rocket.chat/apps-engine/definition/metadata';

const KEY = 'Room';

export interface RoomModel {
	rocketChatRoomId: string;
	teamsThreadId?: string;
	isBridged?: boolean;
	// Set only for rooms linked to a Team channel: the Team's id (channel id lives in teamsThreadId).
	teamsTeamId?: string;
	// Friendly name of the linked Teams chat/channel, captured at link time for display.
	teamsThreadName?: string;
}

export const Room = {
	async persist(
		read: IRead,
		persis: IPersistence,
		rocketChatRoomId: string,
		teamsThreadId?: string,
		extras?: { teamsTeamId?: string; teamsThreadName?: string },
	): Promise<void> {
		const existing = await Room.findByRCRoomId(read, rocketChatRoomId);
		const byRoomId: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatRoomId),
		];
		const data: RoomModel = { ...existing, rocketChatRoomId, teamsThreadId, ...(extras ?? {}) };
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

	async setBridgeActive(persistence: IPersistence, read: IRead, rocketChatRoomId: string, active: boolean): Promise<void> {
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

	// Breaks the room's Teams link: clears thread/team/name, keeps the room record with
	// isBridged=false, and removes the thread-keyed association so inbound lookups stop matching.
	async unlink(read: IRead, persistence: IPersistence, rocketChatRoomId: string): Promise<void> {
		const existing = await Room.findByRCRoomId(read, rocketChatRoomId);
		if (existing?.teamsThreadId) {
			const byThreadId: Array<RocketChatAssociationRecord> = [
				new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
				new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, existing.teamsThreadId),
			];
			await persistence.removeByAssociations(byThreadId);
		}
		const byRoomId: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.MESSAGE, rocketChatRoomId),
		];
		await persistence.updateByAssociations(byRoomId, { rocketChatRoomId, isBridged: false } as RoomModel, true);
	},

	// All room records, deduped by rocketChatRoomId (each is stored under two associations).
	async findAll(read: IRead): Promise<RoomModel[]> {
		const associations: Array<RocketChatAssociationRecord> = [new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY)];
		const results = ((await read.getPersistenceReader().readByAssociations(associations)) ?? []) as RoomModel[];
		const seen = new Set<string>();
		return results.filter((r) => {
			if (!r?.rocketChatRoomId || seen.has(r.rocketChatRoomId)) {
				return false;
			}
			seen.add(r.rocketChatRoomId);
			return true;
		});
	},

	async debugCleanAll(persis: IPersistence): Promise<void> {
		const associations: Array<RocketChatAssociationRecord> = [new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, KEY)];
		await persis.removeByAssociations(associations);
	},
};
