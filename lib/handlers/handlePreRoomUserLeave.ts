import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { UserNotAllowedException } from "@rocket.chat/apps-engine/definition/exceptions";
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

    if (!roomRecord.bridgeUserRocketChatUserId) {
        console.error("No bridge user.");
        throw new UserNotAllowedException();
    }

    const accessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: roomRecord.bridgeUserRocketChatUserId,
        app,
        http,
    });
    if (!accessToken) {
        console.error("No bridge user.");
        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
        throw new UserNotAllowedException();
    }

    const teamsUserId = embeddedLoginUser.teamsUserId;
    if (!teamsUserId) {
        return;
    }

    const threadMemberTeamsUserIds = await listMembersInChatThreadAsync(
        http,
        roomRecord.teamsThreadId,
        accessToken
    );
    if (threadMemberTeamsUserIds.find((id) => id === teamsUserId)) {
        await removeMemberFromChatThreadAsync(
            http,
            roomRecord.teamsThreadId,
            teamsUserId,
            accessToken
        );
    }

    if (
        embeddedLoginUser.teamsUserId === roomRecord.bridgeUserRocketChatUserId
    ) {
        // Clear bridge user if it's been removed
        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
    }
};
