import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { UnsupportedScenarioHintMessageText } from "../Const";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { addMemberToChatThreadAsync } from "../MicrosoftGraphApi";
import { notifyRocketChatUserInRoomAsync } from "../Notifier";
import { Room } from "../PersistHelper";
import { decodeUserOptionValue } from "../UserInterfaceHelper";

export const handleAddTeamsUserContextualBarSubmitAsync = async (options: {
    operator: IUser;
    room: IRoom;
    teamsUserIdsToSave: string[];
    read: IRead;
    persistence: IPersistence;
    http: IHttp;
    modify: IModify;
    app: TeamsBridgeApp;
}): Promise<void> => {
    const {
        app,
        http,
        modify,
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

    const appUser = (await read.getUserReader().getAppUser(app.getID())) as IUser;

    const accessToken = await getUserAccessTokenAsync({ http, app, persistence, read, rocketChatUserId: appUser.id });

    if (!accessToken) {
        await notifyRocketChatUserInRoomAsync(
            UnsupportedScenarioHintMessageText('No valid access token available to add Teams user'),
            appUser,
            operator,
            room,
            read.getNotifier(),
        );
        return;
    }

    const addedNames: string[] = [];
    const alreadyMemberNames: string[] = [];
    const failedNames: string[] = [];

    for (const encodedValue of teamsUserIdsToSave) {
        const { id: teamsUserId, displayName } = decodeUserOptionValue(encodedValue);

        const result = await addMemberToChatThreadAsync(
            http,
            roomRecord.teamsThreadId,
            teamsUserId,
            accessToken,
        );

        if (result.status === 'added') {
            addedNames.push(displayName);
        } else if (result.status === 'already_member') {
            alreadyMemberNames.push(displayName);
        } else {
            failedNames.push(displayName);
        }
    }

    // Permanent message for successfully added users
    if (addedNames.length > 0) {
        const nameList = addedNames.map((n) => `**${n}**`).join(', ');
        const text = addedNames.length === 1
            ? `${nameList} has been added to this channel on MS Teams.`
            : `${nameList} have been added to this channel on MS Teams.`;

        const msg = modify.getCreator().startMessage()
            .setSender(appUser)
            .setRoom(room)
            .setText(text);
        await modify.getCreator().finish(msg);
    }

    // Ephemeral message for users already in the channel
    if (alreadyMemberNames.length > 0) {
        const nameList = alreadyMemberNames.map((n) => `**${n}**`).join(', ');
        const text = alreadyMemberNames.length === 1
            ? `${nameList} is already a member of this channel on MS Teams.`
            : `${nameList} are already members of this channel on MS Teams.`;
        await notifyRocketChatUserInRoomAsync(text, appUser, operator, room, read.getNotifier());
    }

    // Ephemeral message for failed additions
    if (failedNames.length > 0) {
        const nameList = failedNames.map((n) => `**${n}**`).join(', ');
        const text = failedNames.length === 1
            ? `Failed to add ${nameList} to this channel on MS Teams.`
            : `Failed to add ${nameList} to this channel on MS Teams.`;
        await notifyRocketChatUserInRoomAsync(text, appUser, operator, room, read.getNotifier());
    }
};
