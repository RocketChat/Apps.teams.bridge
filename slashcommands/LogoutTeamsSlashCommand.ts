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
    deleteSubscriptionAsync,
    listSubscriptionsAsync,
    revokeUserRefreshTokenAsync,
} from '../lib/MicrosoftGraphApi';
import {
    deleteUserAccessTokenAsync,
    deleteUserAsync,
    retrieveLoginMessageSentStatus,
    retrieveUserAccessTokenAsync,
    saveLoginMessageSentStatus,
} from '../lib/PersistHelper';

export class LogoutTeamsSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-logout-teams';
    public i18nParamsExample: string;
    public i18nDescription: string = 'logout_teams_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence
    ): Promise<void> {
        const notifier = modify.getNotifier();
        const appUser = (await read.getUserReader().getAppUser()) as IUser;
        const sender = context.getSender();
        const currentRoom = context.getRoom();

        // Retrieve existing access token
        const rocketChatUserId = sender.id;
        const userAccessToken = await retrieveUserAccessTokenAsync(
            read,
            persis,
            rocketChatUserId
        );

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
                persistence: persis,
                rocketChatUserId,
                wasSent: false,
            });

            return;
        }

        // Revoke refresh token
        try {
            await revokeUserRefreshTokenAsync(http, userAccessToken);
        } catch (error) {
            console.error(
                `Error during user log out revoking user refresh token. ${error}`
            );
            console.error('This error will be ignored and continue log out.');
        }

        // Delete all subscriptions
        const subscriptionIds = await listSubscriptionsAsync(
            http,
            userAccessToken
        );
        if (subscriptionIds) {
            for (const subscriptionId of subscriptionIds) {
                try {
                    await deleteSubscriptionAsync(
                        http,
                        subscriptionId,
                        userAccessToken
                    );
                } catch (error) {
                    console.error(
                        `Error during delete subscription, will ignore and continue. ${error}`
                    );
                }
            }
        }

        await Promise.all([
            deleteUserAccessTokenAsync(persis, rocketChatUserId),

            // Delete user record
            deleteUserAsync(read, persis, rocketChatUserId),

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
                persistence: persis,
                rocketChatUserId,
                wasSent: false,
            }),
        ]);
    }
}
