import type { IRead } from '@rocket.chat/apps-engine/definition/accessors';

import type { UserModel } from '../persistence';

export const getSenderUser = async ({
	roomRecord,
	fromUserRocketChatUser,
	read,
	fromUserTeamsId,
}: {
	roomRecord: any;
	fromUserRocketChatUser: UserModel | null;
	read: IRead;
	fromUserTeamsId: string;
}) => {
	if (fromUserRocketChatUser) {
		const roomMembers = await read.getRoomReader().getMembers(roomRecord.rocketChatRoomId);
		if (roomMembers?.find((user) => user.id === fromUserRocketChatUser.rocketChatUserId)) {
			return read.getUserReader().getById(fromUserRocketChatUser.rocketChatUserId);
		}
	}

	console.log(`No RC user found for Teams sender ${fromUserTeamsId}, falling back to app bot.`);
	return read.getUserReader().getAppUser();
};
