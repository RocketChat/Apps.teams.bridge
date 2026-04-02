import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { notifyRocketChatUserInRoomAsync } from "../lib/Notifier";
import {
    AddUserRoomTypeInvalidHintMessageText,
    RoomNotBridgedHintMessageText,
} from "../lib/Const";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { openAddTeamsUserContextualBarBlocksAsync } from "../lib/UserInterfaceHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import { Room } from "../lib/PersistHelper";


export class AddUserSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-add-user';
    public i18nParamsExample: string;
    public i18nDescription: string = 'add_user_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    constructor(private app: TeamsBridgeApp) {}

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {

        const currentRoom = context.getRoom();
        const commandSender = context.getSender();
        const appUser = (await read.getUserReader().getAppUser()) as IUser;

        if (currentRoom.type === RoomType.DIRECT_MESSAGE || currentRoom.type === RoomType.CHANNEL) {
            await notifyRocketChatUserInRoomAsync(AddUserRoomTypeInvalidHintMessageText, appUser, commandSender, currentRoom, read.getNotifier());
            return;
        }

        const isBridged = await Room.isBridged(read, currentRoom.id);
        if (!isBridged) {
            await notifyRocketChatUserInRoomAsync(RoomNotBridgedHintMessageText, appUser, commandSender, currentRoom, read.getNotifier());
            return;
        }

        const triggerId = context.getTriggerId() as string;
        await openAddTeamsUserContextualBarBlocksAsync(triggerId, currentRoom, commandSender, appUser, read, modify, http, persis, this.app);
    }
}
