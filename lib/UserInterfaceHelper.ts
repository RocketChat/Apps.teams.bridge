import { IHttp, IModify, IPersistence, IRead, IUIKitSurfaceViewParam } from "@rocket.chat/apps-engine/definition/accessors";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { InputElementDispatchAction, UIKitSurfaceType } from "@rocket.chat/apps-engine/definition/uikit";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import type { TeamsBridgeApp } from "../TeamsBridgeApp";
import { UIActionId, UIElementId, UIElementText } from "./Const";
import { getAppAccessTokenAsync, getUserAccessTokenAsync } from "./AuthHelper";
import { getTeamsChatMembersAsync, searchTeamsUsersAsync } from "./MicrosoftGraphApi";
import type { TeamsChatMember, GetTeamsChatMembersResult } from "./MicrosoftGraphApi";
import { notifyRocketChatUserInRoomAsync } from "./Notifier";
import { Room, UserMapping } from "./PersistHelper";
import type { UserModel } from "./PersistHelper";

interface LoadedUser {
    id: string;
    displayName: string;
}

interface LoadMoreButtonState {
    nextLink: string;
    loadedUsers: LoadedUser[];
    roomId: string;
}

interface ViewMembersButtonState {
    nextLink: string;
    loadedMembers: TeamsChatMember[];
    threadId: string;
}

const encodeButtonState = (state: LoadMoreButtonState): string =>
    Buffer.from(JSON.stringify(state)).toString('base64');

export const decodeButtonState = (value: string): LoadMoreButtonState =>
    JSON.parse(Buffer.from(value, 'base64').toString('utf-8')) as LoadMoreButtonState;

const encodeViewMembersButtonState = (state: ViewMembersButtonState): string =>
    Buffer.from(JSON.stringify(state)).toString('base64');

export const decodeViewMembersButtonState = (value: string): ViewMembersButtonState =>
    JSON.parse(Buffer.from(value, 'base64').toString('utf-8')) as ViewMembersButtonState;

export const encodeUserOptionValue = (id: string, displayName: string): string =>
    Buffer.from(JSON.stringify([id, displayName])).toString('base64');

export const decodeUserOptionValue = (value: string): { id: string; displayName: string } => {
    const [id, displayName] = JSON.parse(Buffer.from(value, 'base64').toString('utf-8')) as [string, string];
    return { id, displayName };
};

export const SearchBlockId = 'TeamsUserSearchBlock';
export const UserSelectBlockId = 'TeamsUserSelectBlock';

export const openAddTeamsUserContextualBarBlocksAsync = async (
    triggerId: string,
    currentRoom: IRoom,
    operator: IUser,
    appUser: IUser,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persistence: IPersistence,
    app: TeamsBridgeApp,
): Promise<void> => {
    const accessToken = await getAppAccessTokenAsync({ read, persistence, http, app });
    if (!accessToken) {
        await notifyRocketChatUserInRoomAsync(
            'No app access token available. Please ensure the app user is logged in to Microsoft Teams.',
            appUser, operator, currentRoom, read.getNotifier(),
        );
        return;
    }

    const { users, nextLink } = await searchTeamsUsersAsync(http, accessToken, {});

    // Build set of Teams user IDs whose RC-registered user is already in this room.
    const members = await read.getRoomReader().getMembers(currentRoom.id);
    const memberTeamsUserIdModels = await Promise.all(
        members.map((m) => UserMapping.findByRCUserId(read, m.id)),
    );
    const memberTeamsUserIdSet = new Set(
        memberTeamsUserIdModels
            .filter((u): u is UserModel => u !== null)
            .map((u) => u.teamsUserId),
    );

    const filteredUsers = users.filter((u) => !memberTeamsUserIdSet.has(u.id));
    const loadedUsers: LoadedUser[] = filteredUsers.map((u) => ({ id: u.id, displayName: u.displayName }));

    const view = createContextualBarBlocks(modify, loadedUsers, currentRoom.id, nextLink, [], '');
    await modify.getUiController().openSurfaceView(view, { triggerId }, operator);
};

export const updateAddTeamsUserContextualBarAsync = async (options: {
    actionId: string;
    value?: string;
    roomId: string;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    modify: IModify;
}): Promise<IUIKitSurfaceViewParam | null> => {
    const { actionId, value, roomId, read, http, persistence, app, modify } = options;

    const accessToken = await getAppAccessTokenAsync({ read, persistence, http, app });
    if (!accessToken) {
        return null;
    }

    if (actionId === UIActionId.TeamsUserSearchInput) {
        // Real-time search on every character change (ON_CHARACTER_ENTERED dispatch).
        // value = the current full text in the search box.
        const query = value ?? '';
        const { users, nextLink } = await searchTeamsUsersAsync(
            http, accessToken, query ? { query } : {},
        );
        const loadedUsers: LoadedUser[] = users.map((u) => ({ id: u.id, displayName: u.displayName }));
        return createContextualBarBlocks(modify, loadedUsers, roomId, nextLink, [], query);
    }

    if (actionId === UIActionId.TeamsUserLoadMore && value) {
        // Append the next Graph page to the accumulated list.
        const { nextLink: prevNextLink, loadedUsers: prevUsers } = decodeButtonState(value);
        const { users: newUsers, nextLink } = await searchTeamsUsersAsync(
            http, accessToken, { pageUrl: prevNextLink },
        );
        const mergedUsers: LoadedUser[] = [
            ...prevUsers,
            ...newUsers.map((u) => ({ id: u.id, displayName: u.displayName })),
        ];
        // Selections are reset because view.state is unavailable in block interactions.
        return createContextualBarBlocks(modify, mergedUsers, roomId, nextLink, [], '');
    }

    return null;
};

export const createContextualBarBlocks = (
    modify: IModify,
    loadedUsers: LoadedUser[],
    roomId: IRoom['id'],
    nextLink?: string,
    selectedUserIds: string[] = [],
    query: string = '',
): IUIKitSurfaceViewParam => {
    const blocks = modify.getCreator().getBlockBuilder();

    blocks.addInputBlock({
        blockId: SearchBlockId,
        element: blocks.newPlainTextInputElement({
            actionId: UIActionId.TeamsUserSearchInput,
            placeholder: blocks.newPlainTextObject(UIElementText.TeamsUserSearchPlaceholder),
            initialValue: query,
            dispatchActionConfig: [InputElementDispatchAction.ON_CHARACTER_ENTERED],
        }),
        label: blocks.newPlainTextObject(UIElementText.TeamsUserSearchLabel),
    });

    if (nextLink) {
        const buttonState: LoadMoreButtonState = { nextLink, loadedUsers, roomId };
        blocks.addActionsBlock({
            elements: [
                blocks.newButtonElement({
                    actionId: UIActionId.TeamsUserLoadMore,
                    text: blocks.newPlainTextObject(UIElementText.TeamsUserLoadMoreButton),
                    value: encodeButtonState(buttonState),
                }),
            ],
        });
    }

    const selectOptions = loadedUsers.map((u) => ({
        text: blocks.newPlainTextObject(u.displayName || u.id),
        value: encodeUserOptionValue(u.id, u.displayName),
    }));

    blocks.addInputBlock({
        blockId: UserSelectBlockId,
        element: blocks.newMultiStaticElement({
            actionId: UIActionId.TeamsUserNameSearch,
            placeholder: blocks.newPlainTextObject(UIElementText.TeamsUserNameSearchPlaceHolder),
            options: selectOptions,
            initialValue: selectedUserIds,
        }),
        label: blocks.newPlainTextObject(UIElementText.TeamsUserNameSearchTitle),
    });

    return {
        id: UIElementId.ContextualBarId,
        title: blocks.newPlainTextObject(UIElementText.ContextualBarTitle),
        type: UIKitSurfaceType.CONTEXTUAL_BAR,
        submit: blocks.newButtonElement({
            actionId: getSubmitActionIdForRoomId(roomId),
            text: blocks.newPlainTextObject(UIElementText.TeamsUsersSaveChangeButton),
        }),
        blocks: blocks.getBlocks(),
    };
};

export const getSubmitActionIdForRoomId = (roomId: IRoom['id']) => `${UIActionId.SaveChanges}--${roomId}`;

export const getRoomIdFromSubmitActionId = (actionId: string) => actionId.trim().split('--').pop();

export const updateViewTeamsMembersContextualBarAsync = async (options: {
    value: string;
    read: IRead;
    http: IHttp;
    persistence: IPersistence;
    app: TeamsBridgeApp;
    modify: IModify;
}): Promise<IUIKitSurfaceViewParam | null> => {
    const { value, read, http, persistence, app, modify } = options;

    const { nextLink: pageUrl, loadedMembers: prevMembers, threadId } = decodeViewMembersButtonState(value);

    const appUser = await read.getUserReader().getAppUser();
    if (!appUser) { return null; }

    const accessToken = await getUserAccessTokenAsync({
        read, persistence, http, app, rocketChatUserId: appUser.id,
    });
    if (!accessToken) { return null; }

    const result = await getTeamsChatMembersAsync(http, threadId, accessToken, { pageUrl });
    if (!result) { return null; }

    const mergedMembers: TeamsChatMember[] = [...prevMembers, ...result.members];
    return createViewMembersContextualBarBlocks(modify, mergedMembers, threadId, result.nextLink);
};

export const openViewTeamsMembersContextualBarAsync = async (
    triggerId: string,
    currentRoom: IRoom,
    operator: IUser,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persistence: IPersistence,
    app: TeamsBridgeApp,
): Promise<void> => {
    const appUser = await read.getUserReader().getAppUser();
    if (!appUser) {
        return;
    }

    const roomRecord = await Room.findByRCRoomId(read, currentRoom.id);
    if (!roomRecord?.teamsThreadId) {
        await notifyRocketChatUserInRoomAsync(
            UIElementText.ViewMembersNoThreadText,
            appUser, operator, currentRoom, read.getNotifier(),
        );
        return;
    }

    const accessToken = await getUserAccessTokenAsync({
        read,
        persistence,
        http,
        app,
        rocketChatUserId: appUser.id,
    });
    if (!accessToken) {
        await notifyRocketChatUserInRoomAsync(
            UIElementText.ViewMembersNoTokenText,
            appUser, operator, currentRoom, read.getNotifier(),
        );
        return;
    }

    const result = await getTeamsChatMembersAsync(http, roomRecord.teamsThreadId, accessToken);
    const view = createViewMembersContextualBarBlocks(
        modify,
        result?.members ?? [],
        roomRecord.teamsThreadId,
        result?.nextLink,
    );
    await modify.getUiController().openSurfaceView(view, { triggerId }, operator);
};

const createViewMembersContextualBarBlocks = (
    modify: IModify,
    members: TeamsChatMember[],
    threadId: string,
    nextLink?: string,
): IUIKitSurfaceViewParam => {
    const blocks = modify.getCreator().getBlockBuilder();

    if (members.length === 0) {
        blocks.addSectionBlock({
            text: blocks.newMarkdownTextObject(UIElementText.ViewMembersEmptyText),
        });
    } else {
        const headerText = `${UIElementText.ViewMembersHeader} *(${members.length} shown)*`;
        blocks.addSectionBlock({
            text: blocks.newMarkdownTextObject(headerText),
        });
        for (const member of members) {
            blocks.addSectionBlock({
                text: blocks.newMarkdownTextObject(`**${member.displayName}**`),
            });
        }
    }

    if (nextLink) {
        const buttonState: ViewMembersButtonState = { nextLink, loadedMembers: members, threadId };
        blocks.addActionsBlock({
            elements: [
                blocks.newButtonElement({
                    actionId: UIActionId.ViewMembersLoadMore,
                    text: blocks.newPlainTextObject(UIElementText.ViewMembersLoadMoreButton),
                    value: encodeViewMembersButtonState(buttonState),
                }),
            ],
        });
    }

    return {
        id: UIElementId.ViewMembersContextualBarId,
        title: blocks.newPlainTextObject(UIElementText.ViewMembersContextualBarTitle),
        type: UIKitSurfaceType.CONTEXTUAL_BAR,
        blocks: blocks.getBlocks(),
    };
};
