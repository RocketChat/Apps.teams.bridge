import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { IUser, UserType } from "@rocket.chat/apps-engine/definition/users";
import { IBotUser } from "@rocket.chat/apps-engine/definition/users/IBotUser";
import { AppSetting } from "../config/Settings";
import { TeamsAppUserNameSurfix } from "./Const";
import {
    getApplicationAccessTokenAsync,
    listTeamsUserProfilesAsync,
} from "./MicrosoftGraphApi";
import {
    persistDummyUserAsync,
    persistTeamsUserProfileAsync,
    retrieveDummyUserByRocketChatUserIdAsync,
    UserModel,
} from "./PersistHelper";

export const syncAllTeamsBotUsersAsync = async (
    http: IHttp,
    read: IRead,
    modify: IModify,
    persis: IPersistence,
    appId: string
): Promise<void> => {
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
    const aadClientSecret = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadClientSecret)
    ).value;

    const response = await getApplicationAccessTokenAsync(
        http,
        aadTenantId,
        aadClientId,
        aadClientSecret
    );
    const appAccessToken = response.accessToken;

    const teamsUserProfiles = await listTeamsUserProfilesAsync(
        http,
        appAccessToken
    );

    for (const profile of teamsUserProfiles) {
        await persistTeamsUserProfileAsync(
            persis,
            profile.displayName,
            profile.givenName,
            profile.surname,
            profile.mail,
            profile.id
        );

        const rocketChatUserId = await createAppUserAsync({
            teamsUserName: profile.displayName,
            read,
            modify,
            appId,
        });
        await persistDummyUserAsync(persis, rocketChatUserId, profile.id);
    }
};

const createAppUserAsync = async ({
    teamsUserName,
    read,
    modify,
    appId,
}: {
    teamsUserName: string;
    read: IRead;
    modify: IModify;
    appId: string;
}): Promise<string> => {
    const rocketChatUserName = `${teamsUserName
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLocaleLowerCase()
        .replace(" ", ".")}.${TeamsAppUserNameSurfix}`;

    const user = await read.getUserReader().getByUsername(rocketChatUserName);
    if (user) {
        return user.id;
    }

    console.log(
        `User ${teamsUserName} + ${rocketChatUserName} does not exist in Rocket Chat. Now create bot for him/her.`
    );

    const modifyCreator = modify.getCreator();

    const data: Partial<IBotUser> = {
        username: rocketChatUserName,
        type: UserType.BOT, // should be UserType.BOT
        isEnabled: true,
        name: teamsUserName,
        roles: ["bot", "Teams Bot"],
        status: "online",
        appId: appId,
    };

    const userBuilder = modifyCreator.startBotUser(data);
    const id = await modifyCreator.finish(userBuilder);

    return id;
};

export const findAllDummyUsersInRocketChatUserListAsync = async (
    read: IRead,
    users: IUser[]
): Promise<UserModel[]> => {
    const result: UserModel[] = [];
    for (const user of users) {
        const userModel = await retrieveDummyUserByRocketChatUserIdAsync(
            read,
            user.id
        );
        if (userModel) {
            result.push(userModel);
        }
    }

    return result;
};
