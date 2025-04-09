import {
    IRead,
    IModify,
    IHttp,
    IPersistence,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    ISlashCommand,
    SlashCommandContext,
} from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import {
    generateHintMessageWithTeamsLoginButton,
    notifyRocketChatUserAsync,
    notifyRocketChatUserInRoomAsync,
} from "../lib/MessageHelper";
import {
    AuthenticationEndpointPath,
    LoginMessageText,
    SubscriberEndpointPath,
} from "../lib/Const";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import {
    retrieveUserAccessTokenAsync,
    retrieveUserByRocketChatUserIdAsync,
} from "../lib/PersistHelper";
import { subscribeToAllMessagesForOneUserAsync } from "../lib/MicrosoftGraphApi";

export class ResubscribeMessages implements ISlashCommand {
    public command: string = "teamsbridge-resubscribe-messages";
    public i18nParamsExample: string;
    public i18nDescription: string =
        "teamsbridge-resubscribe-messages_command_description";

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public constructor(private readonly app: TeamsBridgeApp) {}

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence
    ): Promise<void> {
        const aadTenantId = (
            await read
                .getEnvironmentReader()
                .getSettings()
                .getById(AppSetting.AadTenantId)
        ).value;
        const aadClientId = (
            await read
                .getEnvironmentReader()
                .getSettings()
                .getById(AppSetting.AadClientId)
        ).value;
        const accessors = this.app.getAccessors();
        const authEndpointUrl = await getRocketChatAppEndpointUrl(
            accessors,
            AuthenticationEndpointPath
        );

        const room = context.getRoom();
        const commandSender = context.getSender();
        const loginUrl = getLoginUrl(
            aadTenantId,
            aadClientId,
            authEndpointUrl,
            commandSender.id
        );
        const appUser = (await read.getUserReader().getAppUser()) as IUser;

        const userAccessToken = await retrieveUserAccessTokenAsync(
            read,
            persis,
            commandSender.id
        );
        if (!userAccessToken) {
            const message = generateHintMessageWithTeamsLoginButton(
                loginUrl,
                appUser,
                room,
                LoginMessageText
            );
            await notifyRocketChatUserAsync(
                message,
                commandSender,
                modify.getNotifier()
            );
            return;
        }

        try {
            const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(
                this.app.getAccessors(),
                SubscriberEndpointPath
            );
            const user = await retrieveUserByRocketChatUserIdAsync(
                read,
                commandSender.id
            );
            if (!user) {
                throw new Error(
                    "User not found or the teams user is not synced with Rocket.Chat"
                );
            }
            await subscribeToAllMessagesForOneUserAsync({
                http,
                read,
                persis,
                rocketChatUserId: commandSender.id,
                subscriberEndpointUrl,
                teamsUserId: user.teamsUserId,
                userAccessToken,
                renewIfExists: true,
            });
            const message = `You have been successfully subscribed to messages.`;
            await notifyRocketChatUserInRoomAsync(
                message,
                appUser,
                commandSender,
                room,
                modify.getNotifier()
            );
        } catch (error) {
            this.app
                .getLogger()
                .error(`Failed to subscribe to messages`, error);
            const message = `Failed to subscribe to messages`;
            await notifyRocketChatUserInRoomAsync(
                message,
                appUser,
                commandSender,
                room,
                modify.getNotifier()
            );
        }
    }
}
