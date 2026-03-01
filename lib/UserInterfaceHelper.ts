import { IModify, IRead, IUIKitSurfaceViewParam } from "@rocket.chat/apps-engine/definition/accessors";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { InputElementDispatchAction, IOptionObject, TextObjectType, UIKitSurfaceType } from "@rocket.chat/apps-engine/definition/uikit";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { AddUserNoExistingUsersHintMessageText, UIActionId, UIElementId, UIElementText } from "./Const";
import { notifyRocketChatUserInRoomAsync } from "./MessageHelper";
import { retrieveAllTeamsUserProfilesAsync, retrieveUserByRocketChatUserIdAsync, TeamsUserProfileModel, UserModel } from "./PersistHelper";

export const openAddTeamsUserContextualBarBlocksAsync = async (
    triggerId: string,
    currentRoom: IRoom,
    operator: IUser,
    appUser: IUser,
    read: IRead,
    modify: IModify
) : Promise<void> => {
    const allTeamsUserProfiles = await retrieveAllTeamsUserProfilesAsync(read);
    if (!allTeamsUserProfiles) {
        await notifyRocketChatUserInRoomAsync(AddUserNoExistingUsersHintMessageText, appUser, operator, currentRoom, read.getNotifier());
        return;
    }

    const members = await read.getRoomReader().getMembers(currentRoom.id);
    // Under single-bot arch there are no dummy users. Build a set of Teams user IDs
    // whose registered RC user is already in this room, so they are excluded from the picker.
    const memberTeamsUserIdModels = await Promise.all(
        members.map((m) => retrieveUserByRocketChatUserIdAsync(read, m.id))
    );
    const memberTeamsUserIdSet = new Set(
        memberTeamsUserIdModels
            .filter((u): u is UserModel => u !== null)
            .map((u) => u.teamsUserId)
    );
    const userProfilesNotInRoom = allTeamsUserProfiles.filter(
        (au) => !memberTeamsUserIdSet.has(au.teamsUserId)
    );

    const contextualbarBlocks = createContextualBarBlocks(modify, userProfilesNotInRoom, currentRoom.id);
    await modify.getUiController().openSurfaceView(contextualbarBlocks, { triggerId }, operator);
    return;
};

export const createContextualBarBlocks = (
    modify: IModify,
    userProfilesNotInRoom: TeamsUserProfileModel[],
    roomId: IRoom["id"]
): IUIKitSurfaceViewParam => {
    const blocks = modify.getCreator().getBlockBuilder();
    const selectOptions: IOptionObject[] = [];

    const initialValue: string[] = [];

    for (const userProfile of userProfilesNotInRoom) {
        selectOptions.push({
            text: blocks.newPlainTextObject(userProfile.displayName),
            value: userProfile.teamsUserId,
        });
    }

    const teamsUserNameSearchInput = blocks.newMultiStaticElement({
        actionId: UIActionId.TeamsUserNameSearch,
        placeholder: {
            type: TextObjectType.PLAINTEXT,
            text: UIElementText.TeamsUserNameSearchPlaceHolder,
            emoji: true,
        },
        options: selectOptions,
        initialValue: initialValue,
        dispatchActionConfig: [InputElementDispatchAction.ON_ITEM_SELECTED],
    });

    blocks.addInputBlock({
        element: teamsUserNameSearchInput,
        label: {
            type: TextObjectType.PLAINTEXT,
            text: UIElementText.TeamsUserNameSearchTitle,
        },
    });

    return {
        id: UIElementId.ContextualBarId,
        title: blocks.newPlainTextObject(UIElementText.ContextualBarTitle),
        type: UIKitSurfaceType.CONTEXTUAL_BAR,
        submit: blocks.newButtonElement({
            actionId: getSubmitActionIdForRoomId(roomId),
            text: blocks.newPlainTextObject(
                UIElementText.TeamsUsersSaveChangeButton
            ),
        }),
        blocks: blocks.getBlocks(),
    };
};


export const getSubmitActionIdForRoomId = (roomId: IRoom['id']) => `${UIActionId.SaveChanges}--${roomId}`;

export const getRoomIdFromSubmitActionId = (actionId: string) => actionId.trim().split('--').pop();
