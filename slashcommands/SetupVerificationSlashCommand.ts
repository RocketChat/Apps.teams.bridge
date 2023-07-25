import {
    IRead,
    IModify,
    IHttp,
    IPersistence,
} from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { getApplicationAccessTokenAsync } from "../lib/MicrosoftGraphApi";
import { notifyRocketChatUserInRoomAsync } from "../lib/MessageHelper";
import { AppSetting } from "../config/Settings";
import { persistApplicationAccessTokenAsync } from "../lib/PersistHelper";
import { AppSetupVerificationFailMessageText, AppSetupVerificationPassMessageText } from "../lib/Const";

export class SetupVerificationSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-setup-verification';
    public i18nParamsExample: string;
    public i18nDescription: string = 'setup_verification_slash_command_description';

    // This slash command should only be seen/used by admin user
    public permission?: string | undefined = 'manage-apps';
    public providesPreview: boolean = false;

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {
        const appUser = (await read.getUserReader().getByUsername('microsoftteamsbridge.bot')) as IUser;
        const messageReceiver = context.getSender();
        const room = context.getRoom();

        try {
            const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
            const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
            const aadClientSecret = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientSecret)).value;

            const response = await getApplicationAccessTokenAsync(http, aadTenantId, aadClientId, aadClientSecret);
            await persistApplicationAccessTokenAsync(persis, response.accessToken);

            await notifyRocketChatUserInRoomAsync(AppSetupVerificationPassMessageText, appUser, messageReceiver, room, modify.getNotifier());
        } catch (error) {
            await notifyRocketChatUserInRoomAsync(AppSetupVerificationFailMessageText, appUser, messageReceiver, room, modify.getNotifier());
        }
    }
}
