import {
    IRead,
    IModify,
    IHttp,
    IPersistence,
} from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { getApplicationAccessTokenAsync, listTeamsUserProfilesAsync } from "../lib/MicrosoftGraphApi";
import { notifyRocketChatUserInRoomAsync } from "../lib/MessageHelper";
import { AppSetting } from "../config/Settings";
import { persistDummyUserAsync, persistTeamsUserProfileAsync } from "../lib/PersistHelper";
import { ProvisionTeamsBotUserFailedMessageText, ProvisionTeamsBotUserSucceedMessageText } from "../lib/Const";
import { createAppUserAsync } from "../lib/AppUserHelper";

export class ProvisionTeamsBotUserSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-provision-teams-bot-user';
    public i18nParamsExample: string;
    public i18nDescription: string = 'provision_teams_bot_user_slash_command_description';

    // This slash command should only be seen/used by admin user
    public permission?: string | undefined = 'manage-apps';
    public providesPreview: boolean = false;

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {
        const appUser = (await read.getUserReader().getAppUser()) as IUser;
        const messageReceiver = context.getSender();
        const room = context.getRoom();

        try {
            const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
            const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
            const aadClientSecret = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientSecret)).value;
    
            const response = await getApplicationAccessTokenAsync(http, aadTenantId, aadClientId, aadClientSecret);
            const appAccessToken = response.accessToken;

            const teamsUserProfiles = await listTeamsUserProfilesAsync(http, appAccessToken);

            for (const profile of teamsUserProfiles) {
                await persistTeamsUserProfileAsync(persis, profile.displayName, profile.givenName, profile.surname, profile.mail, profile.id);
                
                const rocketChatUserId = await createAppUserAsync(profile.displayName, profile.mail, read, modify);
                await persistDummyUserAsync(persis, rocketChatUserId, profile.id);
            }

            await notifyRocketChatUserInRoomAsync(ProvisionTeamsBotUserSucceedMessageText, appUser, messageReceiver, room, modify.getNotifier());
        } catch (error) {
            await notifyRocketChatUserInRoomAsync(ProvisionTeamsBotUserFailedMessageText, appUser, messageReceiver, room, modify.getNotifier());
        }
    }
}
