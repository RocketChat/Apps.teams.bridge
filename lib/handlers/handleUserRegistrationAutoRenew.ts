import type { IHttp, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';

import type { TeamsBridgeApp } from '../../TeamsBridgeApp';
import { getUserAccessTokenAsync } from '../AuthHelper';
import { subscribeToAllMessagesForOneUserAsync, subscribeToChannelMessagesAsync } from '../MicrosoftGraphApi';
import { Room, UserMapping, UserRegistration } from '../PersistHelper';

export const handleUserRegistrationAutoRenewAsync = async (options: {
	subscriberEndpointUrl: string;
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
	app: TeamsBridgeApp;
}): Promise<void> => {
	const { http, persistence, read, subscriberEndpointUrl, app } = options;

	const appUser = await read.getUserReader().getAppUser();
	if (!appUser) {
		throw new Error('App user not found');
	}
	const registration = await UserRegistration.findByRCUserId({ read, rocketChatUserId: appUser.id });

	if (registration) {
		try {
			const userAccessToken = await getUserAccessTokenAsync({
				app,
				http,
				persistence,
				read,
				rocketChatUserId: registration.rocketChatUserId,
			});

			if (!userAccessToken) {
				throw new Error(`Failed to get access token for user ${registration.rocketChatUserId}`);
			}

			const user = await UserMapping.findByRCUserId(read, registration.rocketChatUserId);

			if (!user) {
				throw new Error(`User record for user ${registration.rocketChatUserId} not found!`);
			}

			await subscribeToAllMessagesForOneUserAsync({
				read,
				http,
				persis: persistence,
				rocketChatUserId: user.rocketChatUserId,
				subscriberEndpointUrl,
				teamsUserId: user.teamsUserId,
				userAccessToken,
				renewIfExists: true,
				forceRenew: false,
			});
		} catch (error) {
			console.error(`Error during renew registration for user ${registration.rocketChatUserId}. Ignore this error and continue. Error: ${error}`);
		}

		// Renew per-channel subscriptions for channel-linked rooms (chat subscription above
		// does not cover Team channels).
		try {
			const userAccessToken = await getUserAccessTokenAsync({
				app,
				http,
				persistence,
				read,
				rocketChatUserId: registration.rocketChatUserId,
			});
			if (userAccessToken) {
				const rooms = await Room.findAll(read);
				for (const room of rooms) {
					if (room.teamsTeamId && room.teamsThreadId) {
						try {
							await subscribeToChannelMessagesAsync({
								http,
								read,
								persis: persistence,
								rocketChatUserId: registration.rocketChatUserId,
								teamId: room.teamsTeamId,
								channelId: room.teamsThreadId,
								subscriberEndpointUrl,
								userAccessToken,
								renewIfExists: true,
							});
						} catch (error) {
							console.error(`Error renewing channel subscription for room ${room.rocketChatRoomId}: ${error}`);
						}
					}
				}
			}
		} catch (error) {
			console.error(`Error during channel subscription renewal sweep: ${error}`);
		}
	}
};
