import type { IRead, IModify, IHttp, IPersistence } from '@rocket.chat/apps-engine/definition/accessors';
import type { ISlashCommand, SlashCommandContext } from '@rocket.chat/apps-engine/definition/slashcommands';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { BridgeStatusActiveMessageText, BridgeStatusInactiveMessageText } from '../lib/Const';
import { notifyRocketChatUserInRoomAsync } from '../lib/Notifier';
import { Room } from '../lib/PersistHelper';

export class BridgeStatusSlashCommand implements ISlashCommand {
	public command: string = 'teamsbridge-status';

	public i18nParamsExample: string;

	public i18nDescription: string = 'bridge_status_slash_command_description';

	public permission?: string | undefined;

	public providesPreview: boolean = false;

	constructor(private app: TeamsBridgeApp) {}

	public async executor(context: SlashCommandContext, read: IRead, modify: IModify, http: IHttp, persis: IPersistence): Promise<void> {
		const currentRoom = context.getRoom();
		const commandSender = context.getSender();
		const appUser = (await read.getUserReader().getAppUser()) as IUser;

		const isBridged = await Room.isBridged(read, currentRoom.id);

		const message = isBridged ? BridgeStatusActiveMessageText : BridgeStatusInactiveMessageText;

		await notifyRocketChatUserInRoomAsync(message, appUser, commandSender, currentRoom, read.getNotifier());
	}
}
