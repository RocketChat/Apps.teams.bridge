import {
    INotifier,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage, IMessageAction, IMessageAttachment, MessageActionType } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import { AuthenticationEndpointPath, LoginButtonText } from "./Const";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "./UrlHelper";

export const notifyRocketChatUserAsync = async (
    message: IMessage,
    user: IUser,
    notifier: INotifier): Promise<void> => {
    await notifier.notifyUser(user, message);
};

export const notifyRocketChatUserInRoomAsync = async (
    message: string,
    appUser: IUser,
    user: IUser,
    room: IRoom,
    notifier: INotifier): Promise<void> => {
    const messageTemplate: IMessage = {
        text: message,
        sender: appUser,
        room
    };

    await notifyRocketChatUserAsync(messageTemplate, user, notifier);
};

export const generateHintMessageWithTeamsLoginButton = (
    loginUrl: string,
    sender: IUser,
    room: IRoom,
    hintMessageText: string): IMessage => {
    const buttonAction: IMessageAction = {
        type: MessageActionType.BUTTON,
        text: LoginButtonText,
        url: loginUrl,
    };

    const buttonAttachment: IMessageAttachment = {
        actions: [
            buttonAction
        ]
    };

    const message: IMessage = {
        text: hintMessageText,
        sender: sender,
        room,
        attachments: [
            buttonAttachment
        ]
    };

    return message;
};

export const notifyNotLoggedInUserAsync = async (
    read: IRead,
    user: IUser,
    room: IRoom,
    app: TeamsBridgeApp,
    hintMessageText: string
): Promise<void> => {
    const appUser = (await read.getUserReader().getByUsername('microsoftteamsbridge.bot')) as IUser;

    const aadTenantId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadTenantId)
    ).value;
    const aadClientId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadClientId)
    ).value;
    const accessors = app.getAccessors();
    const authEndpointUrl = await getRocketChatAppEndpointUrl(
        accessors,
        AuthenticationEndpointPath
    );
    const loginUrl = getLoginUrl(
        aadTenantId,
        aadClientId,
        authEndpointUrl,
        user.id
    );
    const message = generateHintMessageWithTeamsLoginButton(
        loginUrl,
        appUser,
        room,
        hintMessageText
    );

    await notifyRocketChatUserAsync(message, user, read.getNotifier());
};
