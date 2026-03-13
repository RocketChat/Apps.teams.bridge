import {
    IHttp,
    IModify,
    INotifier,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IMessage, IMessageAction, IMessageAttachment, MessageActionType } from "@rocket.chat/apps-engine/definition/messages";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { TeamsBridgeApp } from "../TeamsBridgeApp";
import {
    AppUserLoginRequiredAdminHintMessageText,
    AppUserLoginRequiredHintMessageText,
    AuthenticationEndpointPath,
    LoginButtonText,
} from "./Const";
import { AppUserLoginNotified } from "./PersistHelper";
import { getLoginUrlAsync, getRocketChatAppEndpointUrl } from "./UrlHelper";

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

/**
 * Notifies all members of a bridged room that the app user is not logged in.
 * Admins receive a login button (URL scoped to the app user's RC ID).
 * Non-admins receive a plain message asking them to contact an admin.
 * Sets the AppUserLoginNotified flag for the room to avoid repeat spam.
 */
export const notifyRoomMembersAppUserNotLoggedInAsync = async (options: {
    read: IRead;
    modify: IModify;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    roomId: string;
}): Promise<void> => {
    const { read, modify, http, persistence, app, roomId } = options;

    const appUser = await read.getUserReader().getAppUser(app.getID());
    if (!appUser) {
        return;
    }

    const room = await read.getRoomReader().getById(roomId);
    if (!room) {
        return;
    }

    const aadTenantId = (
        await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)
    ).value;
    const aadClientId = (
        await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)
    ).value;
    const accessors = app.getAccessors();
    const authEndpointUrl = await getRocketChatAppEndpointUrl(accessors, AuthenticationEndpointPath);

    // Login URL state is the app user's RC ID so the token is stored under the app user
    const loginUrl = await getLoginUrlAsync(persistence, aadTenantId, aadClientId, authEndpointUrl, appUser.id);

    const members = await read.getRoomReader().getMembers(roomId);
    const notifier = modify.getNotifier();

    await Promise.all(
        members
            .filter((member) => member.id !== appUser.id)
            .map(async (member) => {
                const isAdmin = Array.isArray(member.roles) && member.roles.includes('admin');
                if (isAdmin) {
                    const message = generateHintMessageWithTeamsLoginButton(
                        loginUrl,
                        appUser,
                        room,
                        AppUserLoginRequiredAdminHintMessageText,
                    );
                    await notifyRocketChatUserAsync(message, member, notifier);
                } else {
                    await notifyRocketChatUserInRoomAsync(
                        AppUserLoginRequiredHintMessageText,
                        appUser,
                        member,
                        room,
                        notifier,
                    );
                }
            }),
    );

    await AppUserLoginNotified.set(persistence, roomId);
};

export const notifyNotLoggedInUserAsync = async (
    read: IRead,
    persistence: IPersistence,
    user: IUser,
    room: IRoom,
    app: TeamsBridgeApp,
    hintMessageText: string
): Promise<void> => {
    const appUser = (await read.getUserReader().getById(app.getID()));

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
    const loginUrl = await getLoginUrlAsync(
        persistence,
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
