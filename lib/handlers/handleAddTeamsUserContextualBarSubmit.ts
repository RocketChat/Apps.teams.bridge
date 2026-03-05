import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { UnsupportedScenarioHintMessageText } from "../Const";
import { getAppAccessTokenAsync } from "../AuthHelper";
import { addMemberToChatThreadAsync } from "../MicrosoftGraphApi";
import { notifyRocketChatUserInRoomAsync } from "../Notifier";
import { Room } from "../PersistHelper";

export const handleAddTeamsUserContextualBarSubmitAsync = async (options: {
    operator: IUser;
    room: IRoom;
    teamsUserIdsToSave: string[];
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const {
        app,
        http,
        operator,
        persistence,
        read,
        room,
        teamsUserIdsToSave,
    } = options;

    const roomRecord = await Room.findByRCRoomId(read, room.id);
    if (!roomRecord?.teamsThreadId) {
        return;
    }

    // Use the app-level token to add members on Teams side.
    const accessToken = await getAppAccessTokenAsync({ http, app });

    if (!accessToken) {
        const appUser = await read.getUserReader().getAppUser(app.getID()) as IUser;
        await notifyRocketChatUserInRoomAsync(
            UnsupportedScenarioHintMessageText('No valid access token available to add Teams user'),
            appUser,
            operator,
            room,
            read.getNotifier()
        );
        return;
    }

    // Single-bot architecture: iterate directly over Teams user IDs —
    // no dummy RC users to look up or add to the RC room.
    for (const teamsUserId of teamsUserIdsToSave) {
        await addMemberToChatThreadAsync(
            http,
            roomRecord.teamsThreadId,
            teamsUserId,
            accessToken
        );
    }
};
