import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IRoom } from '@rocket.chat/apps-engine/definition/rooms';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { AppSetting } from '../config/Settings';
import { getUserAccessTokenAsync } from './AuthHelper';
import {
	AuthenticationEndpointPath,
	LoginAppUserAlreadyLoggedInMessageText,
	LoginAppUserMessageText,
	LogoutAppUserNoNeedHintMessageText,
	LogoutAppUserSuccessHintMessageText,
} from './Const';
import { deleteAllSubscriptions } from './MicrosoftGraphApi';
import { generateHintMessageWithTeamsLoginButton, notifyRocketChatUserAsync, notifyRocketChatUserInRoomAsync } from './Notifier';
import { AppUserLoginNotified, LoginMessage, UserMapping, UserRegistration } from './PersistHelper';
import { getLoginUrlAsync, getNotificationEndpointUrl, getRocketChatAppEndpointUrl } from './UrlHelper';
import { syncMappingBackupAsync } from './MappingBackup';

interface SessionActionOptions {
	read: IRead;
	modify: IModify;
	http: IHttp;
	persistence: IPersistence;
	app: TeamsBridgeApp;
	operator: IUser;
	room: IRoom;
}

// Returns true when the app bot user currently has a valid Microsoft Teams token.
export const isAppUserLoggedInAsync = async (options: {
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
	app: TeamsBridgeApp;
}): Promise<boolean> => {
	const { read, http, persistence, app } = options;
	const appUser = await read.getUserReader().getAppUser(app.getID());
	if (!appUser) {
		return false;
	}
	const token = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });
	return Boolean(token);
};

// Starts the Teams login flow for the app bot user (posts a "Login Teams" button).
export const performAppUserLoginAsync = async (options: SessionActionOptions): Promise<void> => {
	const { read, modify, http, persistence, app, operator, room } = options;
	const appUser = await read.getUserReader().getAppUser(app.getID());
	if (!appUser) {
		throw new Error('[TeamsBridge] App user not found');
	}

	const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
	const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
	const authEndpointUrl = await getRocketChatAppEndpointUrl(app.getAccessors(), AuthenticationEndpointPath);

	const existingToken = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });
	if (existingToken) {
		await notifyRocketChatUserInRoomAsync(LoginAppUserAlreadyLoggedInMessageText, appUser, operator, room, modify.getNotifier());
		return;
	}

	const loginUrl = await getLoginUrlAsync(persistence, aadTenantId, aadClientId, authEndpointUrl, appUser.id, 'bot');
	const message = generateHintMessageWithTeamsLoginButton(loginUrl, appUser, room, LoginAppUserMessageText);
	await notifyRocketChatUserAsync(message, operator, modify.getNotifier());
};

// Logs the app bot user out of Microsoft Teams and cleans up its stored state.
export const performAppUserLogoutAsync = async (options: SessionActionOptions): Promise<void> => {
	const { read, modify, http, persistence, app, operator, room } = options;
	const notifier = modify.getNotifier();
	const appUser = await read.getUserReader().getAppUser(app.getID());
	if (!appUser) {
		throw new Error('[TeamsBridge] App user not found');
	}

	const appUserToken = await getUserAccessTokenAsync({ read, persistence, http, app, rocketChatUserId: appUser.id });

	if (appUserToken) {
		try {
			await deleteAllSubscriptions(http, appUserToken, await getNotificationEndpointUrl({ appAccessors: app.getAccessors(), rocketChatUserId: appUser.id }));
		} catch (error) {
			app.getLogger().warn(`[TeamsBridge] Failed to delete subscriptions for app user during logout. Continuing cleanup.`, error);
		}
	}

	await Promise.all([
		UserRegistration.delete(persistence, appUser.id),
		UserMapping.delete(read, persistence, appUser.id).then(() => syncMappingBackupAsync({ read, app })),
		LoginMessage.save({ persistence, rocketChatUserId: appUser.id, wasSent: false }),
		AppUserLoginNotified.clearAll(persistence),
		notifyRocketChatUserInRoomAsync(
			appUserToken ? LogoutAppUserSuccessHintMessageText : LogoutAppUserNoNeedHintMessageText,
			appUser,
			operator,
			room,
			notifier,
		),
	]);
};
