import type { IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IMessage } from '@rocket.chat/apps-engine/definition/messages';
import { RocketChatAssociationModel, RocketChatAssociationRecord } from '@rocket.chat/apps-engine/definition/metadata';
import type { IRoom } from '@rocket.chat/apps-engine/definition/rooms';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';
import * as sha256 from 'crypto-js/sha256';

const FOOTPRINT_KEY = 'BridgedMessageFootprint';
const LAST_MESSAGE_KEY = 'BridgedMessage';

export type MessageFootprintInfo = {
	messageFootprint: string;
	createdAt: string;
};

export type LastBridgedMessageInfo = {
	senderId: string;
	text: string;
	roomId: string;
	fileName: string;
};

const simpleHash = (input: string): string => sha256(input).toString();

const calculatePriorityDate = (message: IMessage): Date => {
	const currentDate = new Date();
	let priorityDate = currentDate;
	if (message.updatedAt) {
		const d = new Date(message.updatedAt);
		if (!isNaN(d.getTime())) {
			priorityDate = d;
		}
	} else if (message.createdAt) {
		const d = new Date(message.createdAt);
		if (!isNaN(d.getTime())) {
			priorityDate = d;
		}
	}
	return priorityDate;
};

export const MessageFootprint = {
	generate(message: IMessage, room: IRoom, sender: IUser): string {
		const dateObject = calculatePriorityDate(message);
		dateObject.setSeconds(0);
		dateObject.setMilliseconds(0);
		const timestamp = dateObject.getTime();
		const combinedProperties = `${message.text}${timestamp}` + `${room.id}${room.type}${room.displayName}` + `${sender.id}${sender.username}`;
		return simpleHash(combinedProperties);
	},

	async save(options: { persistence: IPersistence; rocketChatUserId: string; messageFootprint: string }): Promise<string> {
		const { persistence, rocketChatUserId, messageFootprint } = options;
		const associations: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, FOOTPRINT_KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
		];
		return persistence.updateByAssociations(associations, { messageFootprint, createdAt: new Date().toString() }, true);
	},

	async get(options: { read: IRead; rocketChatUserId: string }): Promise<MessageFootprintInfo> {
		const { read, rocketChatUserId } = options;
		const associations: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, FOOTPRINT_KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
		];
		return (await read.getPersistenceReader().readByAssociations(associations)).shift() as MessageFootprintInfo;
	},

	exists(currentFootprint: string, storedInfo: MessageFootprintInfo): boolean {
		return currentFootprint === storedInfo.messageFootprint;
	},

	async getExistenceInfo(message: IMessage, read: IRead): Promise<{ itDoesMessageFootprintExists: boolean; messageFootprint: string } | undefined> {
		try {
			const rocketChatUserId = message.sender.id;
			const storedInfo = await MessageFootprint.get({ rocketChatUserId, read });
			if (!storedInfo) {
				return undefined;
			}
			const messageFootprint = MessageFootprint.generate(message, message.room, message.sender);
			const itDoesMessageFootprintExists = MessageFootprint.exists(messageFootprint, storedInfo);
			return { itDoesMessageFootprintExists, messageFootprint };
		} catch (error) {
			console.error('An error occurred when trying to get message footprint info', error);
			return undefined;
		}
	},

	async getLastBridged(options: { read: IRead; rocketChatUserId: string }): Promise<LastBridgedMessageInfo | null> {
		const { read, rocketChatUserId } = options;
		const associations: Array<RocketChatAssociationRecord> = [
			new RocketChatAssociationRecord(RocketChatAssociationModel.MISC, LAST_MESSAGE_KEY),
			new RocketChatAssociationRecord(RocketChatAssociationModel.USER, rocketChatUserId),
		];
		const result = (await read.getPersistenceReader().readByAssociations(associations)) as unknown as Array<LastBridgedMessageInfo>;
		if (!result) {
			return null;
		}
		return result[0] ?? null;
	},
};
