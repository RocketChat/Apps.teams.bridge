import {
    IRead,
    IModify,
    IHttp,
    IPersistence,
} from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { notifyRocketChatUserInRoomAsync } from "../lib/MessageHelper";
import { ProvisionTeamsBotUserFailedMessageText, ProvisionTeamsBotUserSucceedMessageText } from "../lib/Const";
import { syncAllTeamsBotUsersAsync } from "../lib/AppUserHelper";

export class ProvisionTeamsBotUserSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-provision-teams-bot-user';
    public i18nParamsExample: string;
    public i18nDescription: string = 'provision_teams_bot_user_slash_command_description';

    // This slash command should only be seen/used by admin user
    public permission?: string | undefined = 'manage-apps';
    public providesPreview: boolean = false;

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {
        const appUser = (await read.getUserReader().getAppUser()) as IUser;
        const messageReceiver = context.getSender();
        const room = context.getRoom();

        try {
            await syncAllTeamsBotUsersAsync(http, read, modify, persis);

            await notifyRocketChatUserInRoomAsync(ProvisionTeamsBotUserSucceedMessageText, appUser, messageReceiver, room, modify.getNotifier());
        } catch (error) {
            await notifyRocketChatUserInRoomAsync(ProvisionTeamsBotUserFailedMessageText, appUser, messageReceiver, room, modify.getNotifier());
        }
    }
}
