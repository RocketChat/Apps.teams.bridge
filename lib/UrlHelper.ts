import { randomBytes } from 'crypto';

import type { IAppAccessors, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IApiEndpointMetadata } from '@rocket.chat/apps-engine/definition/api';
import type { IRoom } from '@rocket.chat/apps-engine/definition/rooms';

import { BotUserAuthenticationScopes, getMicrosoftAuthorizeUrl, NormalUserAuthenticationScopes, SubscriberEndpointPath } from './Const';
import { OAuthNonce } from './persistence';
import { AppSetting } from '../config/Settings';

export const getRocketChatAppEndpointUrl = async (appAccessors: IAppAccessors, appEndpointPath: string): Promise<string> => {
	const webhookEndpoint: IApiEndpointMetadata = appAccessors.providedApiEndpoints.find(
		(endpoint) => endpoint.path === appEndpointPath,
	) as IApiEndpointMetadata;
	let siteUrl: string = await appAccessors.environmentReader.getServerSettings().getValueById('Site_Url');

	const proxyUrl = await appAccessors.environmentReader.getSettings().getValueById(AppSetting.ProxyUrl);

	if (proxyUrl && proxyUrl !== '') {
		siteUrl = proxyUrl;
	}

	return new URL(webhookEndpoint.computedPath, siteUrl).toString();
};

export function getNotificationEndpointUrl(params: { appAccessors: IAppAccessors; rocketChatUserId: string }): Promise<string>;
export function getNotificationEndpointUrl(params: { rocketChatUserId: string; subscriberEndpoint: string }): Promise<string>;
export function getNotificationEndpointUrl({
	appAccessors,
	rocketChatUserId,
	subscriberEndpoint,
}: {
	appAccessors?: IAppAccessors;
	rocketChatUserId?: string;
	subscriberEndpoint?: string;
}): Promise<string> {
	if (appAccessors) {
		return getRocketChatAppEndpointUrl(appAccessors, SubscriberEndpointPath).then(
			(subscriberEndpointUrl) => `${subscriberEndpointUrl}?userId=${rocketChatUserId}`,
		);
	}
	if (subscriberEndpoint && rocketChatUserId) {
		return Promise.resolve(`${subscriberEndpoint}?userId=${rocketChatUserId}&hasClientState=1`);
	}
	throw new Error('Invalid parameters');
}

const buildLoginUrl = (
	aadTenantId: string,
	aadClientId: string,
	authEndpointUrl: string,
	userId: string,
	userType: 'normal' | 'bot',
	nonce: string,
): string => {
	const state = Buffer.from(JSON.stringify({ rc_uid: userId, type: userType, nonce })).toString('base64');
	let url = getMicrosoftAuthorizeUrl(aadTenantId);
	url += `?client_id=${aadClientId}`;
	url += '&response_type=code';
	url += `&redirect_uri=${authEndpointUrl}`;
	url += '&response_mode=query';
	url += `&scope=${userType === 'bot' ? BotUserAuthenticationScopes.join('%20') : NormalUserAuthenticationScopes.join('%20')}`;
	url += `&state=${state}`;
	return url;
};

export const getLoginUrlAsync = async (
	persis: IPersistence,
	aadTenantId: string,
	aadClientId: string,
	authEndpointUrl: string,
	userId: string,
	userType: 'normal' | 'bot' = 'normal',
): Promise<string> => {
	const nonce = randomBytes(16).toString('hex');
	await OAuthNonce.persist(persis, userId, nonce);
	return buildLoginUrl(aadTenantId, aadClientId, authEndpointUrl, userId, userType, nonce);
};

export const getRocketChatMessageUrl = async (read: IRead, msgId: string, room: IRoom) => {
	const siteUrl = await read.getEnvironmentReader().getServerSettings().getValueById('Site_Url');
	let roomType = 'channel';
	switch (room.type) {
		case 'p':
			roomType = 'group';
			break;
		case 'd':
			roomType = 'direct';
			break;
		case 'c':
		default:
			roomType = 'channel';
	}
	return `${siteUrl}/${roomType}/${room.id}?msg=${msgId}`;
};
