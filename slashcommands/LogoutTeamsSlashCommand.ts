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
import { notifyRocketChatUserInRoomAsync } from '../lib/MessageHelper';
import {
    deleteAllSubscriptions,
    revokeUserRefreshTokenAsync,
} from '../lib/MicrosoftGraphApi';
import {
    deleteUserAccessTokenAsync,
    deleteUserAsync,
    retrieveLoginMessageSentStatus,
    saveLoginMessageSentStatus,
} from '../lib/PersistHelper';
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
        const appUser = (await read.getUserReader().getByUsername('microsoftteamsbridge.bot')) as IUser;
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

        const wasSent = await retrieveLoginMessageSentStatus({
            read,
            rocketChatUserId,
        });

        if (!userAccessToken) {
            if (wasSent) {
                return;
            }
            // No need to log out
            await notifyRocketChatUserInRoomAsync(
                LogoutNoNeedHintMessageText,
                appUser,
                sender,
                currentRoom,
                notifier
            );

            await saveLoginMessageSentStatus({
                persistence,
                rocketChatUserId,
                wasSent: false,
            });

            return;
        }

        await deleteAllSubscriptions(
            http,
            userAccessToken,
            await getNotificationEndpointUrl({
                appAccessors: this.app.getAccessors(),
                rocketChatUserId,
            })
        );

        // Revoke refresh token
        try {
            await revokeUserRefreshTokenAsync(http, userAccessToken);
        } catch (error) {
            console.error(
                `Error during user log out revoking user refresh token. ${error}`
            );
            console.error('This error will be ignored and continue log out.');
        }

        await Promise.all([
            deleteUserAccessTokenAsync(persistence, rocketChatUserId),

            // Delete user record
            deleteUserAsync(read, persistence, rocketChatUserId),

            // Notify the user
            notifyRocketChatUserInRoomAsync(
                LogoutSuccessHintMessageText,
                appUser,
                sender,
                currentRoom,
                notifier
            ),

            // Set the login message status to false
            saveLoginMessageSentStatus({
                persistence,
                rocketChatUserId,
                wasSent: false,
            }),
        ]);
    }
}
