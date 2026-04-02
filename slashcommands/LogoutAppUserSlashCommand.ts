import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { ISlashCommand, SlashCommandContext } from '@rocket.chat/apps-engine/definition/slashcommands';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { getUserAccessTokenAsync } from '../lib/AuthHelper';
import { LogoutAppUserNoNeedHintMessageText, LogoutAppUserSuccessHintMessageText } from '../lib/Const';
import { deleteAllSubscriptions } from '../lib/MicrosoftGraphApi';
import { notifyRocketChatUserInRoomAsync } from '../lib/Notifier';
import { AppUserLoginNotified, LoginMessage, UserMapping, UserRegistration } from '../lib/PersistHelper';
import { getNotificationEndpointUrl } from '../lib/UrlHelper';

export class LogoutAppUserSlashCommand implements ISlashCommand {
	public command: string = 'teamsbridge-logout-app-user';

	public i18nParamsExample: string = '';

	public i18nDescription: string = 'logout_app_user_slash_command_description';

	public permission: string = 'manage-apps';

	public providesPreview: boolean = false;

	public constructor(private readonly app: TeamsBridgeApp) {}

	public async executor(context: SlashCommandContext, read: IRead, modify: IModify, http: IHttp, persistence: IPersistence): Promise<void> {
		const notifier = modify.getNotifier();
		const appUser = await read.getUserReader().getAppUser(this.app.getID());
		if (!appUser) {
			throw new Error('[TeamsBridge] App user not found');
		}

		const commandSender = context.getSender();
		const currentRoom = context.getRoom();

		if (!commandSender.roles.includes('admin')) {
			await notifyRocketChatUserInRoomAsync('This command is only for admin users.', commandSender, commandSender, currentRoom, notifier);
			return;
		}

		const appUserToken = await getUserAccessTokenAsync({
			read,
			persistence,
			rocketChatUserId: appUser.id,
			app: this.app,
			http,
		});

		// Delete remote subscriptions only when we have a token; cleanup is always performed
		// so stale records don't linger even if the token has already expired or was lost.
		if (appUserToken) {
			try {
				await deleteAllSubscriptions(
					http,
					appUserToken,
					await getNotificationEndpointUrl({
						appAccessors: this.app.getAccessors(),
						rocketChatUserId: appUser.id,
					}),
				);
			} catch (error) {
				this.app.getLogger().warn(`[teamsbridge-logout-app-user] Failed to delete subscriptions for app user. Continuing cleanup.`, error);
			}
		}

		await Promise.all([
			UserRegistration.delete(persistence, appUser.id),
			UserMapping.delete(read, persistence, appUser.id),
			LoginMessage.save({ persistence, rocketChatUserId: appUser.id, wasSent: false }),
			AppUserLoginNotified.clearAll(persistence),
			notifyRocketChatUserInRoomAsync(
				appUserToken ? LogoutAppUserSuccessHintMessageText : LogoutAppUserNoNeedHintMessageText,
				appUser,
				commandSender,
				currentRoom,
				notifier,
			),
		]);
	}
}
