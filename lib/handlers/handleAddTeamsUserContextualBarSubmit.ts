import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { AddUserLoginRequiredHintMessageText } from "../Const";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { addMemberToChatThreadAsync } from "../MicrosoftGraphApi";
import { notifyNotLoggedInUserAsync } from "../Notifier";
import { LoginMessage, Room } from "../PersistHelper";

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

    if (!roomRecord.bridgeUserRocketChatUserId) {
        await notifyNotLoggedInUserAsync(
            read,
            operator,
            room,
            app,
            AddUserLoginRequiredHintMessageText
        );
        return;
    }

    // If there's a thread created in Teams side, update the participants there as well
    const accessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: roomRecord.bridgeUserRocketChatUserId,
        app,
        http,
    });

    if (!accessToken) {
        const wasSent = await LoginMessage.get({
            read,
            rocketChatUserId: operator.id,
        });
        if (!wasSent) {
            await notifyNotLoggedInUserAsync(
                read,
                operator,
                room,
                app,
                AddUserLoginRequiredHintMessageText
            );
            await LoginMessage.save({
                persistence,
                rocketChatUserId: operator.id,
                wasSent: true,
            });
        }
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
