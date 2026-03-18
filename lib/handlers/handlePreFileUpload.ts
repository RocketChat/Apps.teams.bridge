import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IFileUploadContext } from "@rocket.chat/apps-engine/definition/uploads";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { uploadFileToOneDriveAsync } from "../MicrosoftGraphApi";
import { OneDriveFile, Room } from "../PersistHelper";

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
    const appUser = await read.getUserReader().getAppUser();
    if (appUser && senderRocketChatUserId === appUser.id) {
        return;
    }

    if (!await Room.isBridged(read, roomId)) {
        return;
    }

    const roomRecord = await Room.findByRCRoomId(read, roomId);
    if (!roomRecord) {
        throw new Error("No room record find for Teams interop room!");
    }

    // Prefer the sender's own token; fall back to app-level token.
    let userAccessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        rocketChatUserId: senderRocketChatUserId,
        app,
        http,
    });
    if (!userAccessToken && appUser) {
        userAccessToken = await getUserAccessTokenAsync({ http, app, persistence, read, rocketChatUserId: appUser.id });
    }
    if (!userAccessToken) {
        throw new Error("No valid access token available to upload file!");
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
