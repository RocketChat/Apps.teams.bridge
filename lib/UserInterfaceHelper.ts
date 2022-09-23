import { IModify, IRead } from "@rocket.chat/apps-engine/definition/accessors";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";
import { InputElementDispatchAction, IOptionObject, TextObjectType } from "@rocket.chat/apps-engine/definition/uikit";
import { IUIKitContextualBarViewParam } from "@rocket.chat/apps-engine/definition/uikit/UIKitInteractionResponder";
import { IUser } from "@rocket.chat/apps-engine/definition/users";
import { findAllDummyUsersInRocketChatUserListAsync } from "./AppUserHelper";
import { AddUserNoExistingUsersHintMessageText, UIActionId, UIElementId, UIElementText } from "./Const";
import { notifyRocketChatUserInRoomAsync } from "./MessageHelper";
import { retrieveAllTeamsUserProfilesAsync, TeamsUserProfileModel, UserModel } from "./PersistHelper";

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
    const dummyUsers = await findAllDummyUsersInRocketChatUserListAsync(read, members);
    const userProfilesNotInRoom = allTeamsUserProfiles.filter(au => !dummyUsers.find(du => du.teamsUserId == au.teamsUserId));

    const contextualbarBlocks = createContextualBarBlocks(modify, userProfilesNotInRoom);
    await modify.getUiController().openContextualBarView(contextualbarBlocks, { triggerId }, operator);
    return;
};

export const createContextualBarBlocks = (
    modify: IModify,
    userProfilesNotInRoom: TeamsUserProfileModel[]): IUIKitContextualBarViewParam => {
    const blocks = modify.getCreator().getBlockBuilder();
    const selectOptions : IOptionObject[]  = [];

    const initialValue : string[] = [];
    
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
        dispatchActionConfig: [
            InputElementDispatchAction.ON_ITEM_SELECTED
        ],
    });

    blocks.addInputBlock({
        element: teamsUserNameSearchInput,
        label: {
            type: TextObjectType.PLAINTEXT,
            text: UIElementText.TeamsUserNameSearchTitle,
        }
    });

    return {
        id: UIElementId.ContextualBarId,
        title: blocks.newPlainTextObject(UIElementText.ContextualBarTitle),
        submit: blocks.newButtonElement({
            actionId: UIActionId.SaveChanges,
            text: blocks.newPlainTextObject(UIElementText.TeamsUsersSaveChangeButton),
        }),
        blocks: blocks.getBlocks(),
    };
}