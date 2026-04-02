import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { ISlashCommand, SlashCommandContext } from '@rocket.chat/apps-engine/definition/slashcommands';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { AppSetting } from '../config/Settings';
import { getUserAccessTokenAsync } from '../lib/AuthHelper';
import { AuthenticationEndpointPath, LoginAppUserAlreadyLoggedInMessageText, LoginAppUserMessageText } from '../lib/Const';
import { generateHintMessageWithTeamsLoginButton, notifyRocketChatUserAsync, notifyRocketChatUserInRoomAsync } from '../lib/Notifier';
import { getLoginUrlAsync, getRocketChatAppEndpointUrl } from '../lib/UrlHelper';

export class LoginAppUserSlashCommand implements ISlashCommand {
	public command: string = 'teamsbridge-login-app-user';

	public i18nParamsExample: string = '';

	public i18nDescription: string = 'login_app_user_slash_command_description';

	public permission: string = 'manage-apps';

	public providesPreview: boolean = false;

	public constructor(private readonly app: TeamsBridgeApp) {}

	public async executor(context: SlashCommandContext, read: IRead, modify: IModify, http: IHttp, persistence: IPersistence): Promise<void> {
		const appUser = await read.getUserReader().getAppUser(this.app.getID());
		if (!appUser) {
			throw new Error('[TeamsBridge] App user not found');
		}

		const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
		const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;

		const accessors = this.app.getAccessors();
		const authEndpointUrl = await getRocketChatAppEndpointUrl(accessors, AuthenticationEndpointPath);

		const commandSender = context.getSender();
		const room = context.getRoom();

		// Check if the app user already has a valid Teams token
		const existingToken = await getUserAccessTokenAsync({
			read,
			persistence,
			rocketChatUserId: appUser.id,
			app: this.app,
			http,
		});

		if (existingToken) {
			await notifyRocketChatUserInRoomAsync(LoginAppUserAlreadyLoggedInMessageText, appUser, commandSender, room, modify.getNotifier());
			return;
		}

		// Generate a login URL scoped to the app user's RC ID so the OAuth2
		// callback stores the token under the app user's identity
		const loginUrl = await getLoginUrlAsync(persistence, aadTenantId, aadClientId, authEndpointUrl, appUser.id, 'bot');

		const message = generateHintMessageWithTeamsLoginButton(loginUrl, appUser, room, LoginAppUserMessageText);

		await notifyRocketChatUserAsync(message, commandSender, modify.getNotifier());
	}
}
