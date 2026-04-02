import type { IRead, IModify, IHttp, IPersistence } from '@rocket.chat/apps-engine/definition/accessors';
import type { ISlashCommand, SlashCommandContext } from '@rocket.chat/apps-engine/definition/slashcommands';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { AppSetting } from '../config/Settings';
import { getUserAccessTokenAsync } from '../lib/AuthHelper';
import { AuthenticationEndpointPath, LoginMessageText, LoginNoNeedHintMessageText } from '../lib/Const';
import { generateHintMessageWithTeamsLoginButton, notifyRocketChatUserAsync, notifyRocketChatUserInRoomAsync } from '../lib/Notifier';
import { getLoginUrlAsync, getRocketChatAppEndpointUrl } from '../lib/UrlHelper';

export class LoginTeamsSlashCommand implements ISlashCommand {
	public command: string = 'teamsbridge-login-teams';

	public i18nParamsExample: string;

	public i18nDescription: string = 'login_teams_slash_command_description';

	public permission?: string | undefined;

	public providesPreview: boolean = false;

	public constructor(private readonly app: TeamsBridgeApp) {}

	public async executor(context: SlashCommandContext, read: IRead, modify: IModify, http: IHttp, persistence: IPersistence): Promise<void> {
		const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
		const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
		const accessors = this.app.getAccessors();
		const authEndpointUrl = await getRocketChatAppEndpointUrl(accessors, AuthenticationEndpointPath);

		const room = context.getRoom();
		const commandSender = context.getSender();
		const loginUrl = await getLoginUrlAsync(persistence, aadTenantId, aadClientId, authEndpointUrl, commandSender.id, 'normal');
		const appUser = (await read.getUserReader().getAppUser()) as IUser;

		// If the user has already logged, print some other information instead of the login url
		const userAccessToken = await getUserAccessTokenAsync({
			read,
			persistence,
			rocketChatUserId: commandSender.id,
			app: this.app,
			http,
		});
		if (userAccessToken) {
			await notifyRocketChatUserInRoomAsync(LoginNoNeedHintMessageText, appUser, commandSender, room, modify.getNotifier());
			return;
		}

		const message = generateHintMessageWithTeamsLoginButton(loginUrl, appUser, room, LoginMessageText);

		await notifyRocketChatUserAsync(message, commandSender, modify.getNotifier());
	}
}
