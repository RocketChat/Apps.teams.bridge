import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    ISlashCommand,
    SlashCommandContext,
} from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import {
    DeleteTeamsBotUserFailedMessageText,
    DeleteTeamsBotUserSucceedMessageText,
} from "../lib/Const";
import { notifyRocketChatUserInRoomAsync } from "../lib/MessageHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";

export class DeleteTeamsBotUserSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-delete-teams-bot-user';
    public i18nParamsExample: string;
    public i18nDescription: string =
        'delete_teams_bot_user_slash_command_description';

    // This slash command should only be seen/used by admin user
    public permission?: string | undefined = 'manage-apps';
    public providesPreview: boolean = false;


    constructor(private readonly app: TeamsBridgeApp) {}

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence,
    ): Promise<void> {
        const appUser = (await read.getUserReader().getAppUser()) as IUser;
        const messageReceiver = context.getSender();
        const room = context.getRoom();

        try {
            await this.app.deleteAppUsers(modify);

            await notifyRocketChatUserInRoomAsync(
                DeleteTeamsBotUserSucceedMessageText,
                appUser,
                messageReceiver,
                room,
                modify.getNotifier(),
            );
        } catch (error) {
            await notifyRocketChatUserInRoomAsync(
                DeleteTeamsBotUserFailedMessageText,
                appUser,
                messageReceiver,
                room,
                modify.getNotifier(),
            );
        }
    }
}
