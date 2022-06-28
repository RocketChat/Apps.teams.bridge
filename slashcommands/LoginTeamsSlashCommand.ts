import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { sendRocketChatOneOnOneMessageAsync } from "../lib/RocketChatMessageSender";
import { AuthenticationEndpointPath, AuthenticationScopes, MicrosoftLoginBaseUrl } from "../lib/TeamsBridgeConst";
import { getAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";

export class LoginTeamsSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-login-teams';
    public i18nParamsExample: string;
    public i18nDescription: string = 'login_teams_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public constructor(private readonly app: TeamsBridgeApp) {
        this.getLoginUrl = this.getLoginUrl.bind(this);
        this.getLoginMessage = this.getLoginMessage.bind(this);
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
        const authEndpointUrl = await getAppEndpointUrl(accessors, AuthenticationEndpointPath);

        const commandSender = context.getSender();
        const appUser = (await read.getUserReader().getAppUser()) as IUser;

        const loginUrl = this.getLoginUrl(aadTenantId, aadClientId, authEndpointUrl, commandSender.id);
        const message = this.getLoginMessage(loginUrl);

        await sendRocketChatOneOnOneMessageAsync(message, appUser, commandSender, read, modify);
    }

    private getLoginUrl(
        aadTenantId: string,
        aadClientId: string,
        authEndpointUrl: string,
        userId: string): string {
        var url = `${MicrosoftLoginBaseUrl}/${aadTenantId}/oauth2/v2.0/authorize`;
        url += `?client_id=${aadClientId}`;
        url += '&response_type=code';
        url += `&redirect_uri=${encodeURI(authEndpointUrl)}`;
        url += '&response_mode=query';
        url += `&scope=${AuthenticationScopes.join('%20')}`;
        url += `&state=${userId}`;

        return url;
    }

    private getLoginMessage(loginUrl: string): string {
        return 'To start cross platform collaboration, you need to login to Microsoft with your Teams account or guest account. '
            + 'You\'ll be able to keep using Rocket.Chat, but you\'ll also be able to chat with colleagues using Microsoft Teams. '
            + `\n Please click this link to [Login Teams](${loginUrl}).`;
    }
}
