import {
    IAppAccessors,
    IAppInstallationContext,
    IAppUninstallationContext,
    IConfigurationExtend,
    IConfigurationModify,
    IEnvironmentRead,
    IHttp,
    ILogger,
    IModify,
    IPersistence,
    IRead,
  } from '@rocket.chat/apps-engine/definition/accessors';
import {
    ApiSecurity,
    ApiVisibility,
  } from '@rocket.chat/apps-engine/definition/api';
import { App } from '@rocket.chat/apps-engine/definition/App';
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
import { IAppInfo } from '@rocket.chat/apps-engine/definition/metadata';
import {
    IPreRoomUserLeave,
    IRoom,
    IRoomUserLeaveContext,
  } from '@rocket.chat/apps-engine/definition/rooms';
import {
    IJobContext,
    StartupType,
  } from '@rocket.chat/apps-engine/definition/scheduler';
import { ISetting } from '@rocket.chat/apps-engine/definition/settings';
import {
    RoomTypeFilter,
    UIActionButtonContext,
  } from '@rocket.chat/apps-engine/definition/ui';
import {
    IUIKitResponse,
    UIKitActionButtonInteractionContext,
    UIKitViewSubmitInteractionContext,
  } from '@rocket.chat/apps-engine/definition/uikit';
import {
    IFileUploadContext,
    IPreFileUpload,
  } from '@rocket.chat/apps-engine/definition/uploads';
import { UserType } from '@rocket.chat/apps-engine/definition/users';
import { settings } from './config/Settings';
import { AuthenticationEndpoint } from './endpoints/AuthenticationEndpoint';
import { SubscriberEndpoint } from './endpoints/SubscriberEndpoint';
import {
    RegistrationAutoRenewInterval,
    RegistrationAutoRenewSchedulerId,
    SubscriberEndpointPath,
    UIActionId,
    UIElementId,
    WebhookSecretCreationJobId,
  } from './lib/Const';
import {
    handleAddTeamsUserContextualBarSubmitAsync,
    handlePostMessageDeletedAsync,
    handlePostMessageSentAsync,
    handlePostMessageUpdatedAsync,
    handlePreFileUploadAsync,
    handlePreMessageOperationPreventAsync,
    handlePreMessageSentPreventAsync,
    handlePreRoomUserLeaveAsync,
    handleUninstallApp,
    handleUserRegistrationAutoRenewAsync,
  } from './lib/EventHandler';
import { getRocketChatAppEndpointUrl } from './lib/UrlHelper';
import { getRoomIdFromSubmitActionId, openAddTeamsUserContextualBarBlocksAsync } from './lib/UserInterfaceHelper';
import { AddUserSlashCommand } from './slashcommands/AddUserSlashCommand';
import { DeleteTeamsBotUserSlashCommand } from './slashcommands/DeleteTeamsBotUserSlashCommand';
import { LoginTeamsSlashCommand } from './slashcommands/LoginTeamsSlashCommand';
import { LogoutTeamsSlashCommand } from './slashcommands/LogoutTeamsSlashCommand';
import { ProvisionTeamsBotUserSlashCommand } from './slashcommands/ProvisionTeamsBotUserSlashCommand';
import { SetupVerificationSlashCommand } from './slashcommands/SetupVerificationSlashCommand';
import { ResubscribeMessages } from './slashcommands/ResubscriptionMessages';
import { createWebhookSecret, getWebhookSecret } from './lib/PersistHelper';

export class TeamsBridgeApp
    extends App
    implements
      IPreMessageSentPrevent,
      IPostMessageSent,
      IPostMessageUpdated,
      IPreMessageUpdatedPrevent,
      IPostMessageDeleted,
      IPreMessageDeletePrevent,
      IPreFileUpload,
      IPreRoomUserLeave {

    constructor(info: IAppInfo, logger: ILogger, accessors: IAppAccessors) {
      super(info, logger, accessors);
    }

    async getSettingValueById(id: string) {
        return this.getAccessors().environmentReader.getSettings().getValueById(id)
    }

    async onInstall(
        context: IAppInstallationContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify
    ): Promise<void> {
        await createWebhookSecret({ persistence });
    }

    async onEnable(environment: IEnvironmentRead, configurationModify: IConfigurationModify): Promise<boolean> {
        try {
            await configurationModify.scheduler.scheduleOnce({
                id: WebhookSecretCreationJobId,
                when: new Date(),
            });
        } catch (e) {
            this.getLogger().error(e);
        }
        return true;
    }

    async onDisable(configurationModify: IConfigurationModify): Promise<void> {
        await configurationModify.scheduler.cancelAllJobs();
    }

    public async onUninstall(
      context: IAppUninstallationContext,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
      modify: IModify,
    ): Promise<void> {
      return handleUninstallApp(read, http, modify, this)
    }

    public async executePreMessageSentPrevent(
      message: IMessage,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
    ): Promise<boolean> {
        return await handlePreMessageSentPreventAsync({
            app: this,
            message,
            read,
            persistence,
            http,
        });
    }

    public async executePostMessageSent(
      message: IMessage,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
      modify: IModify,
    ): Promise<void> {
        await handlePostMessageSentAsync({
            app: this,
            message,
            read,
            persistence,
            http,
        });
    }

    public async executePreMessageUpdatedPrevent(
      message: IMessage,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
    ): Promise<boolean> {
      return await handlePreMessageOperationPreventAsync({
          app: this,
          message,
          read,
          persistence,
          http,
      });
    }

    public async executePostMessageUpdated(
      message: IMessage,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
      modify: IModify,
    ): Promise<void> {
        await handlePostMessageUpdatedAsync({
            app: this,
            message,
            read,
            persistence,
            http,
        });
    }

    public async executePreMessageDeletePrevent(
      message: IMessage,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
    ): Promise<boolean> {
        return await handlePreMessageOperationPreventAsync({
            app: this,
            message,
            read,
            persistence,
            http,
        });
    }

    public async executePostMessageDeleted(
      message: IMessage,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
      modify: IModify,
      context: IMessageDeleteContext,
    ): Promise<void> {
        await handlePostMessageDeletedAsync({
            app: this,
            message,
            read,
            persistence,
            http,
        });
    }

    public async executePreFileUpload(
      context: IFileUploadContext,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
      modify: IModify,
    ): Promise<void> {
        await handlePreFileUploadAsync({
            app: this,
            context,
            read,
            persistence,
            http
        });
    }

    public async executePreRoomUserLeave(
      context: IRoomUserLeaveContext,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
    ): Promise<void> {
        await handlePreRoomUserLeaveAsync({
            app: this,
            context,
            read,
            persistence,
            http,
        });
    }

    public async executeActionButtonHandler(
      context: UIKitActionButtonInteractionContext,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
      modify: IModify,
    ): Promise<IUIKitResponse> {
      const data = context.getInteractionData();

      if (data.actionId === UIActionId.AddTeamsUserButtonClicked) {
        const appUser = await read.getUserReader().getAppUser();

        if (!appUser) {
          throw new Error('App user not found');
        }

        await openAddTeamsUserContextualBarBlocksAsync(
          data.triggerId,
          data.room,
          data.user,
          appUser,
          read,
          modify,
        );
      }

      return {
        success: true,
      };
    }

    public async executeViewClosedHandler(
          ): Promise<IUIKitResponse> {
      return {
        success: true,
      };
    }

    public async executeViewSubmitHandler(
      context: UIKitViewSubmitInteractionContext,
      read: IRead,
      http: IHttp,
      persistence: IPersistence,
      modify: IModify,
    ): Promise<IUIKitResponse> {
      const { user, view } = context.getInteractionData();

      if (view.id === UIElementId.ContextualBarId) {
        const submitActionId = view.submit?.actionId;

        let currentRoom: IRoom | undefined;
        const roomIdFromActionId = submitActionId && getRoomIdFromSubmitActionId(submitActionId);
        if (roomIdFromActionId) {
            const room = await read
                .getRoomReader()
                .getById(roomIdFromActionId);
            if (room) {
                currentRoom = room;
            }
        }

        let teamsUserIdsToSave: string[] | undefined;

        if (view.state) {
          Object.values(view.state).forEach((item) => {
            Object.entries(item).forEach(([key, value]) => {
              if (key === UIActionId.TeamsUserNameSearch) {
                teamsUserIdsToSave = value as string[] | undefined;
              }
            })
          })
        }

        // Fallback to object property implementation
        if (teamsUserIdsToSave && currentRoom) {
            await handleAddTeamsUserContextualBarSubmitAsync({
                operator: user,
                room: currentRoom,
                teamsUserIdsToSave,
                read,
                modify,
                persistence,
                http,
                app: this,
            });
        }
      }

      return {
        success: true,
      };
    }

    public async deleteAppUsers(modify: IModify): Promise<void> {
        await Promise.all([
          modify.getDeleter().deleteUsers(this.getID(), UserType.APP),
          modify.getDeleter().deleteUsers(this.getID(), UserType.BOT) // To remove old bot users
        ]);
        return;
    }
    protected async extendConfiguration(
      configuration: IConfigurationExtend,
    ): Promise<void> {
      // Register app settings
      await Promise.all(
        settings.map((setting) => configuration.settings.provideSetting(setting)),
      );

      await Promise.all([
        configuration.slashCommands.provideSlashCommand(new SetupVerificationSlashCommand()),
        configuration.slashCommands.provideSlashCommand(new ProvisionTeamsBotUserSlashCommand(this)),
        configuration.slashCommands.provideSlashCommand(new DeleteTeamsBotUserSlashCommand(this)),
        configuration.slashCommands.provideSlashCommand(new LoginTeamsSlashCommand(this)),
        configuration.slashCommands.provideSlashCommand(new LogoutTeamsSlashCommand(this)),
        configuration.slashCommands.provideSlashCommand(new AddUserSlashCommand()),
        configuration.slashCommands.provideSlashCommand(new ResubscribeMessages(this)),
    ]);

      // Register API endpoints
      await configuration.api.provideApi({
        visibility: ApiVisibility.PUBLIC,
        security: ApiSecurity.UNSECURE,
        endpoints: [
          new AuthenticationEndpoint(this),
          new SubscriberEndpoint(this),
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
          ],
        },
      });

      // Config a scheduler for UserAccessToken & Subscription auto renew and start it
      configuration.scheduler.registerProcessors([
        {
          id: RegistrationAutoRenewSchedulerId,
          processor: async (
            jobContext: IJobContext,
            read: IRead,
            modify: IModify,
            http: IHttp,
            persistence: IPersistence,
          ) => {
            try {
              console.log('Start renew registrations!');
              const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(
                this.getAccessors(),
                SubscriberEndpointPath,
              );

              await handleUserRegistrationAutoRenewAsync({
                subscriberEndpointUrl,
                read,
                http,
                persistence,
                app: this,
              });
              console.log('Finish renew registrations!');
            } catch (error) {
              throw new Error(
                `Auto renew registration failed with error: ${error}`,
              );
            }
          },
          startupSetting: {
            type: StartupType.RECURRING,
            interval: RegistrationAutoRenewInterval,
          },
        },
        {
            id: WebhookSecretCreationJobId,
            processor: async (
                jobContext: IJobContext,
                read: IRead,
                modify: IModify,
                http: IHttp,
                persistence: IPersistence,
            ) => {
                try {
                    const webhookSecret = await getWebhookSecret({ persistenceRead: read.getPersistenceReader() });
                    if (!webhookSecret) {
                        this.getLogger().info('Webhook secret is not created. Creating it now.')
                        await createWebhookSecret({ persistence });
                        const subscriberEndpointUrl =
                            await getRocketChatAppEndpointUrl(
                                this.getAccessors(),
                                SubscriberEndpointPath
                            );

                        await handleUserRegistrationAutoRenewAsync({
                            subscriberEndpointUrl,
                            read,
                            http,
                            persistence,
                            app: this,
                        });
                        this.getLogger().info('Webhook secret created and subscriptions were renewed.');
                    }
                } catch (error) {
                    this.getLogger().error(
                        `Webhook secret creation failed with error, Incoming messages may fail to be processed`,
                        error
                    );
                    throw new Error(
                        `Webhook secret creation failed with error: ${error}`,
                    );
                }
            }
        }
      ]);
    }
  }
