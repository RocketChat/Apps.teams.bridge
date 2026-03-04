import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IFileUploadContext } from "@rocket.chat/apps-engine/definition/uploads";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { uploadFileToOneDriveAsync } from "../MicrosoftGraphApi";
import { OneDriveFile, Room, UserMapping } from "../PersistHelper";

export const handlePreFileUploadAsync = async (options: {
    context: IFileUploadContext;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp;
}): Promise<void> => {
    const { context, app, http, persistence, read } = options;
    const senderRocketChatUserId = context.file.userId;
    const roomId = context.file.rid;
    const fileName = context.file.name;
    const fileMIMEType = context.file.type;
    const fileSize = context.file.size;

    if (fileName.startsWith("thumb-")) {
        // TODO: find a better way to not upload the thumb file for image
        return;
    }

    // Skip uploads made by the app bot itself (e.g. inbound relayed files)
    const appUser = await read.getUserReader().getAppUser(app.getID());
    if (appUser && senderRocketChatUserId === appUser.id) {
        return;
    }

    if (!await Room.isBridged(read, roomId)) {
        return;
    }

    // There should be a room record in persist with a bridge user assigned
    const roomRecord = await Room.findByRCRoomId(read, roomId);
    if (!roomRecord) {
        throw new Error("No room record find for Teams interop room!");
    }

    if (!roomRecord.bridgeUserRocketChatUserId) {
        throw new Error("No bridge user assigned to Teams interop room!");
    }

    const bridgeUser = await UserMapping.findByRCUserId(
        read,
        roomRecord.bridgeUserRocketChatUserId
    );
    let userAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: roomRecord.bridgeUserRocketChatUserId,
        app,
        http,
    });
    if (!userAccessToken || !bridgeUser) {
        await Room.persist(
            persistence,
            roomRecord.rocketChatRoomId,
            roomRecord.teamsThreadId,
            undefined
        );
        throw new Error("Invalid bridge user!");
    }

    const senderUserAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: senderRocketChatUserId,
        app,
        http,
    });
    if (senderUserAccessToken) {
        // If file uploader already logged in, make the file uploaded by themselves instead of via the bridge user
        userAccessToken = senderUserAccessToken;
    }

    // Upload the file to One Drive
    const uploadFileResponse = await uploadFileToOneDriveAsync(
        http,
        fileName,
        fileMIMEType,
        fileSize,
        context.content,
        userAccessToken
    );

    // Persist file upload record
    if (uploadFileResponse) {
        await OneDriveFile.persist(
            persistence,
            uploadFileResponse.fileName,
            uploadFileResponse.driveItemId,
        );
    }
};
