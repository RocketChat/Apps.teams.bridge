import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IRoomUserLeaveContext } from "@rocket.chat/apps-engine/definition/rooms";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { listMembersInChatThreadAsync, removeMemberFromChatThreadAsync } from "../MicrosoftGraphApi";
import { Room, UserMapping } from "../PersistHelper";

export const handlePreRoomUserLeaveAsync = async (options: {
    context: IRoomUserLeaveContext;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const { app, context, http, persistence, read } = options;
    const roomId = context.room.id;
    const leavingRocketChatUserId = context.leavingUser.id;

    // When the app bot is removed, pause bridging without touching the Teams thread.
    // Re-adding the bot later will reactivate the same thread.
    const appUser = await read.getUserReader().getAppUser(app.getID());
    if (appUser && leavingRocketChatUserId === appUser.id) {
        await Room.setBridgeActive(persistence, read, roomId, false);
        app.getLogger().info(`[TeamsBridge] Room "${context.room.displayName || roomId}" bridging paused (app user removed).`);
        return;
    }

    const roomRecord = await Room.findByRCRoomId(read, roomId);
    if (!roomRecord || !roomRecord.teamsThreadId) {
        return;
    }

    const embeddedLoginUser = await UserMapping.findByRCUserId(
        read,
        leavingRocketChatUserId
    );

    if (!embeddedLoginUser) {
        return;
    }

    const teamsUserId = embeddedLoginUser.teamsUserId;
    if (!teamsUserId) {
        return;
    }

    // Prefer the leaving user's own delegated token; fall back to app token.
    let accessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: leavingRocketChatUserId,
        app,
        http,
    });
    if (!accessToken) {
        const appUser = await read.getUserReader().getAppUser(app.getID());
        if (appUser) {
            accessToken = await getUserAccessTokenAsync({ http, app, persistence, read, rocketChatUserId: appUser.id });
        }
    }
    if (!accessToken) {
        app.getLogger().warn(`[TeamsBridge] No access token available to remove Teams member ${teamsUserId} from thread ${roomRecord.teamsThreadId}.`);
        return;
    }

    const threadMembers = await listMembersInChatThreadAsync(
        http,
        roomRecord.teamsThreadId,
        accessToken
    );
    const leavingMember = threadMembers.find((member) => member.userId === teamsUserId);
    if (leavingMember) {
        await removeMemberFromChatThreadAsync(
            http,
            roomRecord.teamsThreadId,
            leavingMember.id,
            accessToken
        );
    }
};
