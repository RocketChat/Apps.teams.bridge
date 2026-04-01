import {
    IHttp,
    IPersistence,
    IRead,
    IMessageBuilder,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage } from "@rocket.chat/apps-engine/definition/messages";
import { TeamsBridgeApp } from "../../TeamsBridgeApp";
import { PreventRegistry } from "../PreventRegistry";
import {
    getAvatarUrlForUsername,
    getExtraInfoAndOriginalFileName,
    popExtraInfoAttachment,
} from "../MessageHelper";

export const handlePreMessageSentModifyAsync = async ({
    message,
    builder,
    read,
    http,
    persistence,
    app,
}: {
    message: IMessage;
    builder: IMessageBuilder;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
}): Promise<IMessage> => {
    let extraInfoData = popExtraInfoAttachment(message);
    let targetAttachmentIndex = -1;
    let originalFilename = "";

    if (!extraInfoData || extraInfoData.source !== "ms-teams") {
        message.attachments?.some((att, index) => {
            if (!att.title?.value) {
                return false;
            }

            const { originalFilename: extractedFilename, present, extraInfo } =
                getExtraInfoAndOriginalFileName(att.title.value);

            if (present) {
                extraInfoData = extraInfo;
                targetAttachmentIndex = index;
                originalFilename = extractedFilename;
                return true;
            }
            return false;
        });
    }

    if (extraInfoData?.source === "ms-teams") {
        await PreventRegistry.set(
            persistence,
            `PreventPostMessageHook/${message.id}`,
            true,
        );

        if (typeof extraInfoData.alias === "string") {
            message.alias = extraInfoData.alias;
            message.avatarUrl = await getAvatarUrlForUsername(extraInfoData.alias, read);
        }
    }

    if (targetAttachmentIndex !== -1 && originalFilename) {
        const targetAttachment = message.attachments![targetAttachmentIndex];
        const currentTitle = targetAttachment.title?.value;

        const unmappedFiles = message["_unmappedProperties_"]?.["files"];
        if (Array.isArray(unmappedFiles)) {
            message["_unmappedProperties_"]["files"] = unmappedFiles.map((file: any) =>
                file.name === currentTitle ? { ...file, name: originalFilename } : file
            );
        }

        message.attachments![targetAttachmentIndex] = {
            ...targetAttachment,
            title: {
                ...targetAttachment.title,
                value: originalFilename,
            },
        };

        if (message.file) {
            message.file = { ...message.file, name: originalFilename };
        }
    }

    return message;
};
