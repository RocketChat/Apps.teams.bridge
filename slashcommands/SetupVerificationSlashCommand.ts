import type { IRead, IModify, IHttp, IPersistence } from '@rocket.chat/apps-engine/definition/accessors';
import type { ISlashCommand, SlashCommandContext } from '@rocket.chat/apps-engine/definition/slashcommands';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { AppSetting } from '../config/Settings';
import { getUserAccessTokenAsync } from '../lib/AuthHelper';
import { AppSetupVerificationFailMessageText, AppSetupVerificationPassMessageText, AppUserNotLoggedInSetupVerificationHintText } from '../lib/Const';
import { getApplicationAccessTokenAsync, verifyUserAccessTokenAsync } from '../lib/MicrosoftGraphApi';
import { notifyRocketChatUserInRoomAsync } from '../lib/Notifier';
import { AppToken } from '../lib/PersistHelper';

export class SetupVerificationSlashCommand implements ISlashCommand {
	public command: string = 'teamsbridge-setup-verification';

	public i18nParamsExample: string;

	public i18nDescription: string = 'setup_verification_slash_command_description';

	// This slash command should only be seen/used by admin user
	public permission?: string | undefined = 'manage-apps';

	public providesPreview: boolean = false;

	constructor(private readonly app: TeamsBridgeApp) {}

	public async executor(context: SlashCommandContext, read: IRead, modify: IModify, http: IHttp, persis: IPersistence): Promise<void> {
		const appUser = (await read.getUserReader().getAppUser()) as IUser;
		const messageReceiver = context.getSender();
		const room = context.getRoom();

		try {
			const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
			const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
			const aadClientSecret = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientSecret)).value;

			const response = await getApplicationAccessTokenAsync(http, aadTenantId, aadClientId, aadClientSecret);
			const epochNow = Math.round(Date.now() / 1000);
			await AppToken.persist(persis, response.accessToken, epochNow + response.expiresIn);

			const appUserToken = await getUserAccessTokenAsync({
				read,
				persistence: persis,
				rocketChatUserId: appUser.id,
				http,
				app: this.app,
			});

			if (!appUserToken) {
				await notifyRocketChatUserInRoomAsync(AppUserNotLoggedInSetupVerificationHintText, appUser, messageReceiver, room, modify.getNotifier());
			} else {
				const isTokenValid = await verifyUserAccessTokenAsync(http, appUserToken);
				if (!isTokenValid) {
					await notifyRocketChatUserInRoomAsync(AppUserNotLoggedInSetupVerificationHintText, appUser, messageReceiver, room, modify.getNotifier());
				} else {
					await notifyRocketChatUserInRoomAsync(AppSetupVerificationPassMessageText, appUser, messageReceiver, room, modify.getNotifier());
				}
			}
		} catch (error) {
			await notifyRocketChatUserInRoomAsync(AppSetupVerificationFailMessageText, appUser, messageReceiver, room, modify.getNotifier());
		}
	}
}
