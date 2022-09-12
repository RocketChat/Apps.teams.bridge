import { IModify, IRead } from "@rocket.chat/apps-engine/definition/accessors";
import { IUser, UserType } from "@rocket.chat/apps-engine/definition/users";
import { TeamsAppUserNameSurfix } from "./Const";
import { retrieveDummyUserByRocketChatUserIdAsync, UserModel } from "./PersistHelper";

export const createAppUserAsync = async (teamsUserName: string, email: string, read: IRead, modify: IModify) : Promise<string> => {

    const rocketChatUserName = `${teamsUserName.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase().replace(' ', '.')}.${TeamsAppUserNameSurfix}`;

    const user = await read.getUserReader().getByUsername(rocketChatUserName);
    if (user) {
        return user.id;
    }
    
    console.log(`User ${teamsUserName} + ${rocketChatUserName} does not exist in Rocket Chat. Now create bot for him/her.`);

    const rocketChatEmail = `${email.replace('@', '#')}@${TeamsAppUserNameSurfix}`;

    const modifyCreator = modify.getCreator();

    const data : Partial<IUser> = {
        username: rocketChatUserName,
        emails: [
            {
                address: rocketChatEmail,
                verified: false,
            }
        ],
        type: UserType.APP,
        isEnabled: true,
        name: teamsUserName,
        roles: ["app", "Teams Bot"],
        status: "online",
        utcOffset: -7
    };

    const userBuilder = modifyCreator.startCreateUser(data);
    const id = await modifyCreator.finish(userBuilder);

    return id;
};

export const findAllDummyUsersInRocketChatUserListAsync = async (read: IRead, users: IUser[]) : Promise<UserModel[]> => {
    const result : UserModel[] = [];
    for (const user of users) {
        const userModel = await retrieveDummyUserByRocketChatUserIdAsync(read, user.id);
        if (userModel) {
            result.push(userModel);
        }
    }

    return result;
};
