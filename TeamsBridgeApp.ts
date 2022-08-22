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
import { ISetting } from '@rocket.chat/apps-engine/definition/settings';
import { settings } from './config/Settings';
import { AuthenticationEndpoint } from './endpoints/AuthenticationEndpoint';
import { SubscriberEndpoint } from './endpoints/SubscriberEndpoint';
import {
    handlePostMessageDeletedAsync,
    handlePostMessageSentAsync,
    handlePostMessageUpdatedAsync,
    handlePreMessageOperationPreventAsync,
    handlePreMessageSentPreventAsync
} from './lib/EventHandler';
import { LoginTeamsSlashCommand } from './slashcommands/LoginTeamsSlashCommand';
import { SetupVerificationSlashCommand } from './slashcommands/SetupVerificationSlashCommand';

export class TeamsBridgeApp extends App implements IPreMessageSentPrevent, IPostMessageSent, IPostMessageUpdated, IPreMessageUpdatedPrevent, IPostMessageDeleted, IPreMessageDeletePrevent {
    constructor(info: IAppInfo, logger: ILogger, accessors: IAppAccessors) {
        super(info, logger, accessors);
    }
    
    protected async extendConfiguration(configuration: IConfigurationExtend): Promise<void> {
        // Register app settings
        await Promise.all(settings.map((setting) => configuration.settings.provideSetting(setting)));
        
        // Register slash commands
        await configuration.slashCommands.provideSlashCommand(new SetupVerificationSlashCommand());
        await configuration.slashCommands.provideSlashCommand(new LoginTeamsSlashCommand(this));
        
        // Register API endpoints
        await configuration.api.provideApi({
            visibility: ApiVisibility.PUBLIC,
            security: ApiSecurity.UNSECURE,
            endpoints: [
                new AuthenticationEndpoint(this),
                new SubscriberEndpoint(this)
            ],
        });
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
        return await handlePreMessageSentPreventAsync(message, read);
    }

    public async executePostMessageSent(
        message: IMessage,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify): Promise<void> {
        await handlePostMessageSentAsync(message, read, http, persistence, modify, this);
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
}
