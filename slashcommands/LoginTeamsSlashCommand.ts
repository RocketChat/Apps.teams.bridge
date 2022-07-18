import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { nofityRocketChatUserAsync } from "../lib/Messages";
import { AuthenticationEndpointPath, AuthenticationScopes, getMicrosoftAuthorizeUrl, LoginButtonText, LoginMessageText } from "../lib/Const";
import { getRocketChatAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import { IMessage, IMessageAction, IMessageAttachment, MessageActionType } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";

export class LoginTeamsSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-login-teams';
    public i18nParamsExample: string;
    public i18nDescription: string = 'login_teams_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public constructor(private readonly app: TeamsBridgeApp) {
        this.getLoginUrl = this.getLoginUrl.bind(this);
        this.getLoginMessageWithButton = this.getLoginMessageWithButton.bind(this);
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
        const loginUrl = this.getLoginUrl(aadTenantId, aadClientId, authEndpointUrl, commandSender.id);
        const appUser = (await read.getUserReader().getAppUser()) as IUser;

        // TODO: check whether current user has already logged in
        // If the user has already logged, print some other information instead of the login url
        const message = this.getLoginMessageWithButton(loginUrl, appUser, room);

        await nofityRocketChatUserAsync(message, commandSender, modify);
    }

    private getLoginMessageWithButton(loginUrl: string, appUser: IUser, room: IRoom) : IMessage {
        const buttonAction: IMessageAction = {
            type: MessageActionType.BUTTON,
            text: LoginButtonText,
            url: loginUrl,
        };

        const buttonAttachment: IMessageAttachment = {
            actions: [
                buttonAction
            ]
        };

        const message: IMessage = {
            text: LoginMessageText,
            sender: appUser,
            room,
            attachments: [
                buttonAttachment
            ]
        };

        return message;
    }

    private getLoginUrl(
        aadTenantId: string,
        aadClientId: string,
        authEndpointUrl: string,
        userId: string): string {
        let url = getMicrosoftAuthorizeUrl(aadTenantId);
        url += `?client_id=${aadClientId}`;
        url += '&response_type=code';
        url += `&redirect_uri=${authEndpointUrl}`;
        url += '&response_mode=query';
        url += `&scope=${AuthenticationScopes.join('%20')}`;
        url += `&state=${userId}`;

        return url;
    }
}
