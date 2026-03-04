import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { getUserAccessTokenAsync } from "../AuthHelper";
import { combineRocketChatMessagesToTeamsMessage, isBridgedMessageFormat } from "../MessageHelper";
import {
    deleteTextMessageInChatThreadAsync,
    updateTextMessageInChatThreadAsync,
} from "../MicrosoftGraphApi";
import { MessageMapping, UploadMapping, UserMapping } from "../PersistHelper";
import type { UploadMappingModel } from "../PersistHelper";
import { PreventRegistry } from "../PreventRegistry";

export const handlePostMessageDeletedAsync = async (options: {
    message: IMessage;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp;
}): Promise<void> => {
    const { message, read, persistence, app, http } = options;
    if (
        await PreventRegistry.capture(
            persistence,
            `PreventPostMessageDeleteHook/${message.id}`
        )
    ) {
        // Prevent duplicate processing
        return;
    }

    const msgId = message.id;
    if (!msgId) {
        return;
    }

    // --- Step 1: Resolve mappings (message ↔ upload ↔ teams) ---
    const {
        messageIdMapping,
        uploadMappings,
        currentUploadMapping,
        mainMessage,
    } = await resolveMappings(read, { ...message, id: msgId });

    if (
        !messageIdMapping &&
        !currentUploadMapping &&
        uploadMappings.length === 0
    ) {
        return;
    }

    // --- Step 2: Ensure sender info (user + access token) ---
    const { senderUser, accessToken } = await ensureSenderInfo({
        senderId: message.sender.id,
        read,
        persistence,
        app,
        http,
    });

    if (!senderUser || !accessToken) {
        return;
    }

    // --- Step 3: Clean up mappings in persistence ---
    if (currentUploadMapping) {
        await UploadMapping.delete({
            persistence,
            rocketchatUploadId: currentUploadMapping.rocketchatUploadId,
            teamsMessageId: currentUploadMapping.teamsMessageId,
        });
    }

    if (messageIdMapping?.rocketChatMessageId === msgId) {
        await MessageMapping.delete({ persistence, ...messageIdMapping });
    }

    // --- Step 4: Prepare Teams update ---
    const teamsIds = {
        messageId:
            messageIdMapping?.teamsMessageId ||
            currentUploadMapping?.teamsMessageId,
        threadId:
            messageIdMapping?.teamsThreadId ||
            currentUploadMapping?.teamsThreadId,
    };

    if (!teamsIds.messageId || !teamsIds.threadId) {
        return;
    }

    const deletedIds = {
        messages: new Set([msgId]),
        uploads: message.file?._id ? new Set<string>([message.file._id]) : new Set<string>(),
    };

    const isBridge = isBridgedMessageFormat(mainMessage?.text || "");
    const { text, shouldDeleteTeamsMessage, attachments } =
        await combineRocketChatMessagesToTeamsMessage({
            read,
            messages: mainMessage ? [mainMessage] : [],
            messageIdMapping: {
                rocketChatMessageId: msgId,
                teamsMessageId: teamsIds.messageId,
                teamsThreadId: teamsIds.threadId,
            },
            deletedMessages: deletedIds.messages,
            deletedUploads: deletedIds.uploads,
            forceBridgedMessage: isBridge,
            originalSenderName: isBridge
                ? message.sender.name || message.sender.username
                : undefined,
            uploadMappings,
            http,
            accessToken,
        });

    // --- Step 5: Execute Teams update/delete ---
    if (shouldDeleteTeamsMessage) {
        await PreventRegistry.set(
            persistence,
            `PreventPostMessageDeleteHook/${message.id}`
        );
        await deleteTextMessageInChatThreadAsync(
            http,
            senderUser.teamsUserId,
            teamsIds.messageId,
            teamsIds.threadId,
            accessToken
        );
    } else {
        await PreventRegistry.set(
            persistence,
            `PreventPostMessageUpdateHook/${message.id}`
        );
        await updateTextMessageInChatThreadAsync({
            http,
            textMessage: text,
            messageType: "html",
            messageId: teamsIds.messageId,
            threadId: teamsIds.threadId,
            userAccessToken: accessToken,
            attachments,
        });
    }
};

async function resolveMappings(read: IRead, message: IMessage & { id: string }) {
    let mainMessage: IMessage | null = null;
    let messageIdMapping =
        await MessageMapping.findByRCMessageId(
            read,
            message.id
        );

    let uploadMappings: UploadMappingModel[] = [];
    if (message.file?._id) {
        uploadMappings =
            await UploadMapping.findAllByRCUploadId(
                read,
                message.file._id
            );
    }

    const currentUploadMapping = uploadMappings.find(
        (u) => u.rocketchatUploadId === message.file?._id
    );

    if (currentUploadMapping && !messageIdMapping) {
        messageIdMapping = await MessageMapping.findByTeamsMessageId(
            read,
            currentUploadMapping.teamsMessageId
        );
        if (messageIdMapping) {
            mainMessage =
                (await read
                    .getMessageReader()
                    .getById(messageIdMapping.rocketChatMessageId)) || null;
        }
    } else if (!currentUploadMapping && messageIdMapping) {
        uploadMappings = await UploadMapping.findByTeamsMessageId(
            read,
            messageIdMapping.teamsMessageId
        );
        mainMessage = message;
    }

    return {
        messageIdMapping,
        uploadMappings,
        currentUploadMapping,
        mainMessage,
    };
}

async function ensureSenderInfo({
    senderId,
    read,
    persistence,
    app,
    http,
}: {
    senderId: string;
    read: IRead;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    http: IHttp;
}) {
    const [accessToken, senderUser] = await Promise.all([
        getUserAccessTokenAsync({
            read,
            persistence,
            rocketChatUserId: senderId,
            app,
            http,
        }),
        UserMapping.findByRCUserId(read, senderId),
    ]);

    return { senderUser, accessToken };
}
