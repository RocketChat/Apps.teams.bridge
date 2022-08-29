import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { generateHintMessageWithTeamsLoginButton, notifyRocketChatUserAsync } from "../lib/MessageHelper";
import { AuthenticationEndpointPath, LoginMessageText } from "../lib/Const";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";

export class LoginTeamsSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-login-teams';
    public i18nParamsExample: string;
    public i18nDescription: string = 'login_teams_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public constructor(private readonly app: TeamsBridgeApp) {
    }

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {
        const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
        const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
        const accessors = this.app.getAccessors();
        const authEndpointUrl = await getRocketChatAppEndpointUrl(accessors, AuthenticationEndpointPath);

        const room = context.getRoom();
        const commandSender = context.getSender();
        const loginUrl = getLoginUrl(aadTenantId, aadClientId, authEndpointUrl, commandSender.id);
        const appUser = (await read.getUserReader().getAppUser()) as IUser;

        // TODO: check whether current user has already logged in
        // If the user has already logged, print some other information instead of the login url

        const message = generateHintMessageWithTeamsLoginButton(loginUrl, appUser, room, LoginMessageText);

        await notifyRocketChatUserAsync(message, commandSender, modify.getNotifier());
    }
}
