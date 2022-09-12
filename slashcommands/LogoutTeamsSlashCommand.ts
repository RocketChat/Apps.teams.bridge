import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { notifyRocketChatUserInRoomAsync } from "../lib/MessageHelper";
import { LogoutNoNeedHintMessageText, LogoutSuccessHintMessageText } from "../lib/Const";
import { deleteUserAccessTokenAsync, retrieveUserAccessTokenAsync } from "../lib/PersistHelper";
import { deleteSubscriptionAsync, listSubscriptionsAsync, revokeUserRefreshTokenAsync } from "../lib/MicrosoftGraphApi";

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
        persis: IPersistence): Promise<void> {
        const notifier = modify.getNotifier();
        const appUser = (await read.getUserReader().getAppUser()) as IUser;
        const sender = context.getSender();
        const currentRoom = context.getRoom();

        // Retrieve existing access token
        const rocketChatUserId = sender.id;
        const userAccessToken = await retrieveUserAccessTokenAsync(read, rocketChatUserId);

        if (!userAccessToken) {
            // No need to log out
            await notifyRocketChatUserInRoomAsync(LogoutNoNeedHintMessageText, appUser, sender, currentRoom, notifier);
            return;
        }
        
        // Revoke refresh token
        try {
            await revokeUserRefreshTokenAsync(http, userAccessToken);
        } catch (error) {
            console.error(`Error during user log out revoking user refresh token. ${error}`);
            console.error('This error will be ignored and continue log out.')
        }

        // Delete all subscriptions
        const subscriptionIds = await listSubscriptionsAsync(http, userAccessToken);
        if (subscriptionIds) {
            for (const subscriptionId of subscriptionIds) {
                try {
                    await deleteSubscriptionAsync(http, subscriptionId, userAccessToken);
                } catch (error) {
                    console.error(`Error during delete subscription, will ignore and continue. ${error}`);
                }
            }
        }
        
        // Delete access token record
        await deleteUserAccessTokenAsync(persis, rocketChatUserId);

        // Notify the user
        await notifyRocketChatUserInRoomAsync(LogoutSuccessHintMessageText, appUser, sender, currentRoom, notifier);
    }
}
