import { IRead } from "@rocket.chat/apps-engine/definition/accessors";
import { UserModel } from "../persistence";

export const getSenderUser = async ({
        roomRecord,
        fromUserRocketChatUser,
        read,
        fromUserTeamsId,
    }: {
        roomRecord: any,
        fromUserRocketChatUser: UserModel | null,
        read: IRead,
        fromUserTeamsId: string,
}) => {
    if (fromUserRocketChatUser) {
        const roomMembers = await read.getRoomReader().getMembers(roomRecord.rocketChatRoomId);
        if (roomMembers && roomMembers.find((user) => user.id === fromUserRocketChatUser.rocketChatUserId)) {
            return read.getUserReader().getById(fromUserRocketChatUser.rocketChatUserId);
        }
    }

    // Under single-bot arch there are no dummy users. Fall back to the app bot
    // so the message is still relayed to RC under the bridge bot identity.
    console.log(
        `No RC user found for Teams sender ${fromUserTeamsId}, falling back to app bot.`
    );
    return read.getUserReader().getAppUser();
}
