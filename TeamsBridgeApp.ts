import {
    IAppAccessors,
    IConfigurationExtend,
    ILogger,
} from '@rocket.chat/apps-engine/definition/accessors';
import { ApiSecurity, ApiVisibility } from '@rocket.chat/apps-engine/definition/api';
import { App } from '@rocket.chat/apps-engine/definition/App';
import { IAppInfo } from '@rocket.chat/apps-engine/definition/metadata';
import { settings } from './config/Settings';
import { AuthenticationEndpoint } from './endpoints/AuthenticationEndpoint';
import { LoginTeamsSlashCommand } from './slashcommands/LoginTeamsSlashCommand';
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
        await configuration.slashCommands.provideSlashCommand(new LoginTeamsSlashCommand(this));
        
        // Register API endpoints
        await configuration.api.provideApi({
            visibility: ApiVisibility.PUBLIC,
            security: ApiSecurity.UNSECURE,
            endpoints: [ new AuthenticationEndpoint(this) ],
        });
    }
}
