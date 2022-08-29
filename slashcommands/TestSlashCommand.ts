import { IRead, IModify, IHttp, IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import { ISlashCommand, SlashCommandContext } from "@rocket.chat/apps-engine/definition/slashcommands";
import { IUser, UserStatusConnection, UserType } from "@rocket.chat/apps-engine/definition/users";
import { AppSetting } from "../config/Settings";
import { generateHintMessageWithTeamsLoginButton, notifyRocketChatUserAsync } from "../lib/MessageHelper";
import { AuthenticationEndpointPath, LoginMessageText } from "../lib/Const";
import { getLoginUrl, getRocketChatAppEndpointUrl } from "../lib/UrlHelper";
import { TeamsBridgeApp } from "../TeamsBridgeApp";

export class TestSlashCommand implements ISlashCommand {
    public command: string = 'teamsbridge-test';
    public i18nParamsExample: string;
    public i18nDescription: string = 'login_teams_slash_command_description';

    public permission?: string | undefined;
    public providesPreview: boolean = false;

    public constructor(private readonly app: TeamsBridgeApp) {
    }

    public async executor(
        context: SlashCommandContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<void> {
        /*
        const userReader = read.getUserReader();
        const user = await userReader.getById("v4ECCH3pTAE6nBXyJ");
        console.log(user); //*/
        //*
        const modifyCreator = modify.getCreator();

        const data : Partial<IUser> = {
            username: "alex.wilber",
            type: UserType.APP,
            isEnabled: true,
            name: "Alex Wilber",
            roles: ["app"],
            status: "online",
            utcOffset: -7
        };

        const userBuilder = modifyCreator.startCreateUser(data);
        const id = await modifyCreator.finish(userBuilder);
        console.log("User Created!");
        console.log(id);
        //*/
    }
}
