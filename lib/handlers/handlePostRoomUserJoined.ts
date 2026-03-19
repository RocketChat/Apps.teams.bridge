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
import { AppUserLoginNotified, Room, UserMapping } from "../PersistHelper";
import { AppUserAddedToRoomMessageText } from "../Const";
import { addMemberToChatThreadAsync } from "../MicrosoftGraphApi";
import { PreventRegistry } from "../PreventRegistry";

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
    if (!appUser) {
        return;
    }

    if (joiningUser.id !== appUser.id) {
        const roomRecord = await Room.findByRCRoomId(read, room.id);
        if (!roomRecord || !roomRecord.isBridged || !roomRecord.teamsThreadId) {
            return;
        }

        const embeddedLoginUser = await UserMapping.findByRCUserId(read, joiningUser.id);
        if (!embeddedLoginUser?.teamsUserId) {
            return;
        }

        const appUserToken = await getUserAccessTokenAsync({ http, app, persistence, read, rocketChatUserId: appUser.id });
        if (!appUserToken) {
            app.getLogger().warn(`[TeamsBridge] No app user access token available to add Teams member ${embeddedLoginUser.teamsUserId} to thread ${roomRecord.teamsThreadId}.`);
            return;
        }

        await PreventRegistry.set(persistence, `member-add:${roomRecord.teamsThreadId}:${embeddedLoginUser.teamsUserId}`);

        await addMemberToChatThreadAsync(
            http,
            roomRecord.teamsThreadId,
            embeddedLoginUser.teamsUserId,
            appUserToken
        );

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
    } else if (inviter) {
        await notifyRocketChatUserInRoomAsync(
            AppUserAddedToRoomMessageText,
            appUser,
            inviter,
            room,
            modify.getNotifier(),
        );
    }
};
