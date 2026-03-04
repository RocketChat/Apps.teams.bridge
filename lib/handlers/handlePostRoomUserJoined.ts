import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IRoomUserJoinedContext } from "@rocket.chat/apps-engine/definition/rooms";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { Room } from "../PersistHelper";

export const handlePostRoomUserJoinedAsync = async (options: {
    context: IRoomUserJoinedContext;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const { context, read, persistence, app } = options;
    const { joiningUser, room } = context;

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
};
