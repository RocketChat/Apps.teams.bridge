import {
    IRead,
    IModify,
    IHttp,
    IPersistence,
} from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { getApplicationAccessTokenAsync } from "../lib/MicrosoftGraphApi";
import { sendRocketChatOneOnOneMessageAsync } from "../lib/RocketChatMessageSender";

export class SetupVerificationSlashCommand implements ISlashCommand {
    private verificationPassMessage: string = 'TeamsBridge app setup verification PASSED!';
    private verificationFailMessage: string =
        'TeamsBridge app setup verification FAILED! Please check trouble shooting guide for further actions.';

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
        const result = await getApplicationAccessTokenAsync(read, http, persis);

        const appUser = (await read.getUserReader().getAppUser()) as IUser;
        const messageReceiver = context.getSender();
        if (result) {
            await sendRocketChatOneOnOneMessageAsync(this.verificationPassMessage, appUser, messageReceiver, read, modify);
        } else {
            await sendRocketChatOneOnOneMessageAsync(this.verificationFailMessage, appUser, messageReceiver, read, modify);
        }
    }
}
