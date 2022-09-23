import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { generateHintMessageWithTeamsLoginButton, notifyRocketChatUserAsync } from "../lib/MessageHelper";
import { AuthenticationEndpointPath, LoginMessageText, RegistrationAutoRenewSchedulerId } from "../lib/Const";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import { listSubscriptionsAsync, renewSubscriptionAsync, renewUserAccessTokenAsync } from "../lib/MicrosoftGraphApi";
import { persistUserAccessTokenAsync, retrieveAllUserRegistrationsAsync, retrieveUserAccessTokenAsync, retrieveUserRefreshTokenAsync } from "../lib/PersistHelper";

export class TestSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-test';
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

    }
}
