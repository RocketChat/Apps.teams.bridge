import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";
import { notifyRocketChatUserInRoomAsync } from "../lib/MessageHelper";
import {
    AddUserRoomTypeInvalidHintMessageText,
} from "../lib/Const";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { openAddTeamsUserContextualBarBlocksAsync } from "../lib/UserInterfaceHelper";


export class AddUserSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-add-user';
    public i18nParamsExample: string;
    public i18nDescription: string = 'add_user_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {
        const [subcommand] = context.getArguments();

        const currentRoom = context.getRoom();
        const commandSender = context.getSender();
        const appUser = (await read.getUserReader().getByUsername('microsoftteamsbridge.bot')) as IUser;

        if (currentRoom.type === RoomType.DIRECT_MESSAGE || currentRoom.type === RoomType.CHANNEL) {
            await notifyRocketChatUserInRoomAsync(AddUserRoomTypeInvalidHintMessageText, appUser, commandSender, currentRoom, read.getNotifier());
            return;
        }

        if (!subcommand) {
            // If no subcommand is provided, open search Teams User ContextualBar
            const triggerId = context.getTriggerId() as string;
            await openAddTeamsUserContextualBarBlocksAsync(triggerId, currentRoom, commandSender, appUser, read, modify);

            return;
        }

        // Under single-bot arch there are no dummy users to add via username.
        // Direct the user to use the contextual bar instead.
        const triggerId = context.getTriggerId() as string;
        await openAddTeamsUserContextualBarBlocksAsync(triggerId, currentRoom, commandSender, appUser, read, modify);
    }
}
