import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IRoomUserJoinedContext } from "@rocket.chat/apps-engine/definition/rooms";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { notifyRocketChatUserInRoomAsync, notifyRoomMembersAppUserNotLoggedInAsync } from "../Notifier";
import { AppUserLoginNotified, Room } from "../PersistHelper";
import { AppUserAddedToRoomMessageText } from "../Const";

export const handlePostRoomUserJoinedAsync = async (options: {
    context: IRoomUserJoinedContext;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    modify: IModify;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const { context, read, persistence, modify, app, http } = options;
    const { joiningUser, room, inviter } = context;

    const appUser = await read.getUserReader().getAppUser(app.getID());
    if (!appUser || joiningUser.id !== appUser.id) {
        return;
    }

    // setBridgeActive preserves any existing teamsThreadId /
    // bridgeUserRocketChatUserId, so re-adding the bot reuses the same thread
    await Room.setBridgeActive(persistence, read, room.id, true);

    app.getLogger().info(
        `[TeamsBridge] Room "${room.displayName || room.id}" is now an active bridge room ` +
        `(app user added by ${context.inviter?.username ?? 'unknown'}).`
    );

    const appUserToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: appUser.id,
        app,
        http,
    });
    if (!appUserToken) {
        await notifyRoomMembersAppUserNotLoggedInAsync({
            read,
            modify,
            http,
            persistence,
            app,
            roomId: room.id,
        });
    } else if(inviter){
        await notifyRocketChatUserInRoomAsync(
            AppUserAddedToRoomMessageText,
            appUser,
            inviter,
            room,
            modify.getNotifier(),
        );
    }
};
