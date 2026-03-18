import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
    ISlashCommand,
    SlashCommandContext,
} from '@rocket.chat/apps-engine/definition/slashcommands';
import { IUser } from '@rocket.chat/apps-engine/definition/users';
import {
    LogoutNoNeedHintMessageText,
    LogoutSuccessHintMessageText,
} from '../lib/Const';
import { notifyRocketChatUserInRoomAsync } from '../lib/Notifier';
import { deleteAllSubscriptions } from '../lib/MicrosoftGraphApi';
import { LoginMessage, UserMapping, UserRegistration } from '../lib/PersistHelper';
import { getNotificationEndpointUrl } from '../lib/UrlHelper';
import { TeamsBridgeApp } from '../TeamsBridgeApp';
import { getUserAccessTokenAsync } from '../lib/AuthHelper';

export class LogoutTeamsSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-logout-teams';
    public i18nParamsExample: string;
    public i18nDescription: string = 'logout_teams_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public constructor(private readonly app: TeamsBridgeApp) {}

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persistence: IPersistence
    ): Promise<void> {
        const notifier = modify.getNotifier();
        const appUser = await read.getUserReader().getAppUser();
        const sender = context.getSender();
        const currentRoom = context.getRoom();

        // Retrieve existing access token
        const rocketChatUserId = sender.id;
        const userAccessToken = await getUserAccessTokenAsync({
            read,
            persistence,
            rocketChatUserId,
            app: this.app,
            http,
        });

        // Delete remote subscriptions only when we have a token; cleanup is always performed
        // so stale records don't linger even if the token has already expired or was lost.
        if (userAccessToken) {
            try {
                await deleteAllSubscriptions(
                    http,
                    userAccessToken,
                    await getNotificationEndpointUrl({
                        appAccessors: this.app.getAccessors(),
                        rocketChatUserId,
                    })
                );
            } catch (error) {
                this.app.getLogger().warn(`[teamsbridge-logout-teams] Failed to delete subscriptions for user ${rocketChatUserId}. Continuing cleanup.`, error);
            }
        }

        await Promise.all([
            UserRegistration.delete(persistence, rocketChatUserId),
            UserMapping.delete(read, persistence, rocketChatUserId),
            appUser ? notifyRocketChatUserInRoomAsync(
                userAccessToken ? LogoutSuccessHintMessageText : LogoutNoNeedHintMessageText,
                appUser,
                sender,
                currentRoom,
                notifier
            ) : Promise.resolve(),
            LoginMessage.save({
                persistence,
                rocketChatUserId,
                wasSent: false,
            }),
        ]);
    }
}
