import {
    IAppAccessors,
    IConfigurationExtend,
    IConfigurationModify,
    IHttp,
    ILogger,
    IModify,
    IPersistence,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import { ApiSecurity, ApiVisibility } from '@rocket.chat/apps-engine/definition/api';
import { App } from '@rocket.chat/apps-engine/definition/App';
import { 
    IMessage,
    IMessageDeleteContext,
    IPostMessageDeleted,
    IPostMessageSent,
    IPostMessageUpdated,
    IPreMessageDeletePrevent,
    IPreMessageSentPrevent,
    IPreMessageUpdatedPrevent
} from '@rocket.chat/apps-engine/definition/messages';
import { IAppInfo } from '@rocket.chat/apps-engine/definition/metadata';
import { IPreRoomUserLeave, IRoom, IRoomUserLeaveContext } from '@rocket.chat/apps-engine/definition/rooms';
import { IJobContext, StartupType } from '@rocket.chat/apps-engine/definition/scheduler';
import { ISetting } from '@rocket.chat/apps-engine/definition/settings';
import { RoomTypeFilter, UIActionButtonContext } from '@rocket.chat/apps-engine/definition/ui';
import {
    IUIKitResponse,
    UIKitActionButtonInteractionContext,
    UIKitBlockInteractionContext,
    UIKitViewCloseInteractionContext,
    UIKitViewSubmitInteractionContext
} from '@rocket.chat/apps-engine/definition/uikit';
import { IFileUploadContext, IPreFileUpload } from '@rocket.chat/apps-engine/definition/uploads';
import { IUser } from '@rocket.chat/apps-engine/definition/users';
import { settings } from './config/Settings';
import { AuthenticationEndpoint } from './endpoints/AuthenticationEndpoint';
import { SubscriberEndpoint } from './endpoints/SubscriberEndpoint';
import { RegistrationAutoRenewInterval, RegistrationAutoRenewSchedulerId, SubscriberEndpointPath, UIActionId, UIElementId } from './lib/Const';
import {
    handleAddTeamsUserContextualBarSubmitAsync,
    handlePostMessageDeletedAsync,
    handlePostMessageSentAsync,
    handlePostMessageUpdatedAsync,
    handlePreFileUploadAsync,
    handlePreMessageOperationPreventAsync,
    handlePreMessageSentPreventAsync,
    handlePreRoomUserLeaveAsync,
    handleUserRegistrationAutoRenewAsync
} from './lib/EventHandler';
import { getRocketChatAppEndpointUrl } from './lib/UrlHelper';
import { openAddTeamsUserContextualBarBlocksAsync } from './lib/UserInterfaceHelper';
import { AddUserSlashCommand } from './slashcommands/AddUserSlashCommand';
import { LoginTeamsSlashCommand } from './slashcommands/LoginTeamsSlashCommand';
import { LogoutTeamsSlashCommand } from './slashcommands/LogoutTeamsSlashCommand';
import { ProvisionTeamsBotUserSlashCommand } from './slashcommands/ProvisionTeamsBotUserSlashCommand';
import { SetupVerificationSlashCommand } from './slashcommands/SetupVerificationSlashCommand';
import { TestSlashCommand } from './slashcommands/TestSlashCommand';

export class TeamsBridgeApp extends App implements IPreMessageSentPrevent, IPostMessageSent, IPostMessageUpdated, IPreMessageUpdatedPrevent, IPostMessageDeleted, IPreMessageDeletePrevent, IPreFileUpload, IPreRoomUserLeave {
    private selectedTeamsUserIds: string[];
    private changeTeamsUserMemberRoom: IRoom | undefined;
    
    constructor(info: IAppInfo, logger: ILogger, accessors: IAppAccessors) {
        super(info, logger, accessors);
    }
    
    protected async extendConfiguration(configuration: IConfigurationExtend): Promise<void> {
        // Register app settings
        await Promise.all(settings.map((setting) => configuration.settings.provideSetting(setting)));
        
        // Register slash commands
        await configuration.slashCommands.provideSlashCommand(new SetupVerificationSlashCommand());
        await configuration.slashCommands.provideSlashCommand(new ProvisionTeamsBotUserSlashCommand());
        await configuration.slashCommands.provideSlashCommand(new LoginTeamsSlashCommand(this));
        await configuration.slashCommands.provideSlashCommand(new LogoutTeamsSlashCommand());
        await configuration.slashCommands.provideSlashCommand(new AddUserSlashCommand());
        await configuration.slashCommands.provideSlashCommand(new TestSlashCommand(this));
        
        // Register API endpoints
        await configuration.api.provideApi({
            visibility: ApiVisibility.PUBLIC,
            security: ApiSecurity.UNSECURE,
            endpoints: [
                new AuthenticationEndpoint(this),
                new SubscriberEndpoint(this)
            ],
        });

        // Config context menu item
        configuration.ui.registerButton({
            actionId: UIActionId.AddTeamsUserButtonClicked,
            labelI18n: 'action_button_label_add_teams_user',
            context: UIActionButtonContext.ROOM_ACTION,
            when: {
                roomTypes: [
                    RoomTypeFilter.PRIVATE_DISCUSSION, 
                    RoomTypeFilter.PRIVATE_CHANNEL, 
                    RoomTypeFilter.PRIVATE_TEAM,
                ]
            }
        });

        // Config a scheduler for UserAccessToken & Subscription auto renew and start it
        configuration.scheduler.registerProcessors([
            {
                id: RegistrationAutoRenewSchedulerId,
                processor: async (jobContext: IJobContext, read: IRead, modify: IModify, http: IHttp, persis: IPersistence) => {
                    try {
                        console.log("Start renew registrations!")
                        const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(this.getAccessors(), SubscriberEndpointPath);

                        await handleUserRegistrationAutoRenewAsync(subscriberEndpointUrl, read, modify, http, persis);
                        console.log("Finish renew registrations!")
                    } catch (error) {
                        throw new Error(`Auto renew registration failed with error: ${error}`);
                    }
                },
                startupSetting: {
                    type: StartupType.RECURRING,
                    interval: RegistrationAutoRenewInterval,
                }
            }
        ]);
    }
    
    public async onSettingUpdated(
        setting: ISetting,
        configurationModify: IConfigurationModify,
        read: IRead,
        http: IHttp): Promise<void> {
        console.log(`onSettingUpdated for setting ${setting.id} with new value ${setting.value}`);
    }

    public async executePreMessageSentPrevent(
        message: IMessage,
        read: IRead,
        http: IHttp,
        persistence: IPersistence): Promise<boolean> {
        return await handlePreMessageSentPreventAsync(message, read, persistence, this);
    }

    public async executePostMessageSent(
        message: IMessage,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify): Promise<void> {
        await handlePostMessageSentAsync(message, read, http, persistence);
    }
    
    public async executePreMessageUpdatedPrevent(
        message: IMessage,
        read: IRead,
        http: IHttp,
        persistence: IPersistence): Promise<boolean> {
        return await handlePreMessageOperationPreventAsync(message, read);
    }

    public async executePostMessageUpdated(
        message: IMessage,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify): Promise<void> {
        await handlePostMessageUpdatedAsync(message, read, http);
    }

    public async executePreMessageDeletePrevent(
        message: IMessage,
        read: IRead,
        http: IHttp,
        persistence: IPersistence): Promise<boolean> {
        return await handlePreMessageOperationPreventAsync(message, read);
    }

    public async executePostMessageDeleted(
        message: IMessage,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify,
        context: IMessageDeleteContext): Promise<void> {
        await handlePostMessageDeletedAsync(message, read, http);
    }

    public async executePreFileUpload(
        context: IFileUploadContext,
        read: IRead,
        http: IHttp,
        persis: IPersistence,
        modify: IModify): Promise<void> {
        await handlePreFileUploadAsync(context, read, http, persis, modify);
    }

    public async executePreRoomUserLeave(
        context: IRoomUserLeaveContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence): Promise<void> {
        await handlePreRoomUserLeaveAsync(context, read, http, persistence, this);
    }

    public async executeActionButtonHandler(
        context: UIKitActionButtonInteractionContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify
    ): Promise<IUIKitResponse> {
        const data = context.getInteractionData();

        if (data.actionId === UIActionId.AddTeamsUserButtonClicked) {        
            const appUser = (await read.getUserReader().getAppUser()) as IUser;
            await openAddTeamsUserContextualBarBlocksAsync(
                data.triggerId,
                data.room,
                data.user,
                appUser,
                read,
                modify);
        }

        return {
            success: true,
        };
    }

    public async executeBlockActionHandler(
        context: UIKitBlockInteractionContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify) : Promise<IUIKitResponse> {
        const data = context.getInteractionData();
        if (data.actionId === UIActionId.TeamsUserNameSearch) {
            this.changeTeamsUserMemberRoom = data.room;
            this.selectedTeamsUserIds = data.value as any as string[];
        }

        return {
            success: true,
        };
    }

    public async executeViewClosedHandler(
        context: UIKitViewCloseInteractionContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify
    ) : Promise<IUIKitResponse> {
        return {
            success: true,
        };
    }

    public async executeViewSubmitHandler(
        context: UIKitViewSubmitInteractionContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify): Promise<IUIKitResponse> {
        const data = context.getInteractionData();
        if (data.view.id === UIElementId.ContextualBarId) {
            if (this.changeTeamsUserMemberRoom) {
                await handleAddTeamsUserContextualBarSubmitAsync(
                    context.getInteractionData().user,
                    this.changeTeamsUserMemberRoom,
                    this.selectedTeamsUserIds,
                    read,
                    modify,
                    persistence,
                    http,
                    this);
            }
        }

        return {
            success: true,
        };
    }
}
