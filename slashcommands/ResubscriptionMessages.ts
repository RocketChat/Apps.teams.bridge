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
} from "../lib/Notifier";
import {
    AuthenticationEndpointPath,
    LoginAppUserMessageText,
    SubscriberEndpointPath,
} from "../lib/Const";
import { getLoginUrlAsync, getRocketChatAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import { UserMapping } from "../lib/PersistHelper";
import { subscribeToAllMessagesForOneUserAsync } from "../lib/MicrosoftGraphApi";
import { getUserAccessTokenAsync } from "../lib/AuthHelper";

export class ResubscribeMessages implements ISlashCommand {
    public command: string = "teamsbridge-resubscribe-messages";
    public i18nParamsExample: string;
    public i18nDescription: string =
        "teamsbridge-resubscribe-messages_command_description";

    public permission?: string | undefined = "manage-apps";
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

        if (!commandSender.roles.includes("admin")) {
            const message = "This command is only for admin users.";
            await notifyRocketChatUserInRoomAsync(
                message,
                commandSender,
                commandSender,
                room,
                modify.getNotifier()
            );
            return;
        }
        const appUser = (await read.getUserReader().getAppUser()) as IUser;

        const userAccessToken = await getUserAccessTokenAsync({
            read,
            persistence: persis,
            rocketChatUserId: appUser.id,
            app: this.app,
            http,
        });
        if (!userAccessToken) {
            const loginUrl = await getLoginUrlAsync(
                persis,
                aadTenantId,
                aadClientId,
                authEndpointUrl,
                appUser.id,
                "bot",
            );
            const message = generateHintMessageWithTeamsLoginButton(
                loginUrl,
                appUser,
                room,
                LoginAppUserMessageText,
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
            const userMapping = await UserMapping.findByRCUserId(
                read,
                appUser.id
            );
            if (!userMapping) {
                throw new Error(
                    "User not found or the teams user is not synced with Rocket.Chat"
                );
            }
            await subscribeToAllMessagesForOneUserAsync({
                http,
                read,
                persis,
                rocketChatUserId: userMapping.rocketChatUserId,
                subscriberEndpointUrl,
                teamsUserId: userMapping.teamsUserId,
                userAccessToken,
                renewIfExists: true,
                forceRenew: true,
            });
            const message = `The bot has been successfully subscribed to messages.`;
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
