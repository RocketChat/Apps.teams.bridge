import {
    IAppAccessors,
    IConfigurationExtend,
    IConfigurationModify,
    IHttp,
    ILogger,
    IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import { ApiSecurity, ApiVisibility } from '@rocket.chat/apps-engine/definition/api';
import { App } from '@rocket.chat/apps-engine/definition/App';
import { IAppInfo } from '@rocket.chat/apps-engine/definition/metadata';
import { ISetting } from '@rocket.chat/apps-engine/definition/settings';
import { AppSetting, settings } from './config/Settings';
import { AuthenticationEndpoint } from './endpoints/AuthenticationEndpoint';
import { SetupVerificationSlashCommand } from './slashcommands/SetupVerificationSlashCommand';

export class TeamsBridgeApp extends App {
    constructor(info: IAppInfo, logger: ILogger, accessors: IAppAccessors) {
        super(info, logger, accessors);
    }
    
    protected async extendConfiguration(configuration: IConfigurationExtend): Promise<void> {
        // Register app settings
        await Promise.all(settings.map((setting) => configuration.settings.provideSetting(setting)));
        
        // Register slash commands
        await configuration.slashCommands.provideSlashCommand(new SetupVerificationSlashCommand());
        
        // Register API endpoints
        await configuration.api.provideApi({
            visibility: ApiVisibility.PUBLIC,
            security: ApiSecurity.UNSECURE,
            endpoints: [ new AuthenticationEndpoint(this) ],
        });
    }
    
    public async onSettingUpdated(
        setting: ISetting,
        configurationModify: IConfigurationModify,
        read: IRead,
        http: IHttp): Promise<void> {
        const aadTenantId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadTenantId)).value;
        const aadClientId = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientId)).value;
        const aadClientSecret = (await read.getEnvironmentReader().getSettings().getById(AppSetting.AadClientSecret)).value;

        console.log(`onSettingUpdated for setting ${setting.id} with new value ${setting.value}`);
    }
}
