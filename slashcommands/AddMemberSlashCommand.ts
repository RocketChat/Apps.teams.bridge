import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser, UserStatusConnection, UserType } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { generateHintMessageWithTeamsLoginButton, notifyRocketChatUserAsync } from "../lib/MessageHelper";
import { AuthenticationEndpointPath, LoginMessageText } from "../lib/Const";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import { RoomType } from "@rocket.chat/apps-engine/definition/rooms";

export class AddMemberSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-add-member';
    public i18nParamsExample: string;
    public i18nDescription: string = 'login_teams_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {
        const [subcommand] = context.getArguments();

        if (!subcommand) {
            return;
        }

        const currentRoom = context.getRoom();

        if (currentRoom.type === RoomType.DIRECT_MESSAGE) {
            return;
        }

        const commandSender = context.getSender();

        const updater = modify.getUpdater();
        const roomBuilder = await updater.room(currentRoom.id, commandSender);

        const userToAdd = await read.getUserReader().getByUsername(subcommand);

        roomBuilder.addMemberToBeAddedByUsername(userToAdd.username);
        await updater.finish(roomBuilder);
    }
}
