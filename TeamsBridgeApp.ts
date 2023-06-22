import {
  ApiSecurity,
  ApiVisibility,
  IAppAccessors,
  IAppInfo,
  IAppUninstallationContext,
  IConfigurationExtend,
  IConfigurationModify,
  IHttp,
  ILogger,
  IModify,
  IPersistence,
  IRead,
} from '@rocket.chat/apps-engine/definition/accessors';
import {
  IMessage,
  IMessageDeleteContext,
  IPostMessageDeleted,
  IPostMessageSent,
  IPostMessageUpdated,
  IPreMessageDeletePrevent,
  IPreMessageSentPrevent,
  IPreMessageUpdatedPrevent,
} from '@rocket.chat/apps-engine/definition/messages';
import {
  IPreRoomUserLeave,
  IRoom,
  IRoomUserLeaveContext,
} from '@rocket.chat/apps-engine/definition/rooms';
import {
  IFileUploadContext,
  IPreFileUpload,
} from '@rocket.chat/apps-engine/definition/uploads';
import { IUser, UserType } from '@rocket.chat/apps-engine/definition/users';
import { ISetting } from '@rocket.chat/apps-engine/definition/settings';
import { settings } from './config/Settings';
import { AuthenticationEndpoint } from './endpoints/AuthenticationEndpoint';
import { SubscriberEndpoint } from './endpoints/SubscriberEndpoint';
import {
  RegistrationAutoRenewInterval,
  RegistrationAutoRenewSchedulerId,
  SubscriberEndpointPath,
  UIActionId,
  UIElementId,
} from './lib/Const';
import * as EventHandler from './lib/EventHandler';
import { getRocketChatAppEndpointUrl } from './lib/UrlHelper';
import { openAddTeamsUserContextualBarBlocksAsync } from './lib/UserInterfaceHelper';
import {
  AddUserSlashCommand,
  DeleteTeamsBotUserSlashCommand,
  LoginTeamsSlashCommand,
  LogoutTeamsSlashCommand,
  ProvisionTeamsBotUserSlashCommand,
  SetupVerificationSlashCommand,
} from './slashcommands';

export class TeamsBridgeApp extends App implements
  IPreMessageSentPrevent,
  IPostMessageSent,
  IPostMessageUpdated,
  IPreMessageUpdatedPrevent,
  IPostMessageDeleted,
  IPreMessageDeletePrevent,
  IPreFileUpload,
  IPreRoomUserLeave {
  private selectedTeamsUserIds: string[];
  private changeTeamsUserMemberRoom?: IRoom;

  constructor(info: IAppInfo, logger: ILogger, accessors: IAppAccessors) {
    super(info, logger, accessors);

    this.selectedTeamsUserIds = [];
  }

  public async onEnable(configurationModify: IConfigurationModify): Promise<void> {
    await Promise.all([
      configurationModify.slashCommands.provideSlashCommand(new AddUserSlashCommand()),
      configurationModify.slashCommands.provideSlashCommand(new DeleteTeamsBotUserSlashCommand()),
      configurationModify.slashCommands.provideSlashCommand(new LoginTeamsSlashCommand()),
      configurationModify.slashCommands.provideSlashCommand(new LogoutTeamsSlashCommand()),
      configurationModify.slashCommands.provideSlashCommand(new ProvisionTeamsBotUserSlashCommand()),
      configurationModify.slashCommands.provideSlashCommand(new SetupVerificationSlashCommand()),
      configurationModify.api.registerApiEndpoint(new AuthenticationEndpoint(this)),
      configurationModify.api.registerApiEndpoint(new SubscriberEndpoint(this)),
      configurationModify.settings.provideSetting(settings),
    ]);
  }

  public async onAppUninstalled(context: IAppUninstallationContext, read: IRead): Promise<void> {
    // Clean up app data and subscriptions related to the uninstalled app
  }

  // Implement other event handlers and methods

  // ...
}

export function extendConfiguration(configuration: IConfigurationExtend): void {
  configuration.api.provideApi({
    visibility: ApiVisibility.PRIVATE,
    security: ApiSecurity.UNSECURE,
    endpoints: [getRocketChatAppEndpointUrl(SubscriberEndpointPath)],
  });

  configuration.scheduler.registerProcessors([
    {
      id: RegistrationAutoRenewSchedulerId,
      processor: async () => {
        // Logic to renew user access tokens and subscriptions
      },
      interval: RegistrationAutoRenewInterval,
    },
  ]);

  configuration.settings.provideSetting(settings);
}

export async function createApp(
  info: IAppInfo,
  logger: ILogger,
  accessors: IAppAccessors,
): Promise<TeamsBridgeApp> {
  const app = new TeamsBridgeApp(info, logger, accessors);
  return app;
}
