import {
    IAppAccessors,
    IAppInstallationContext,
    IAppUninstallationContext,
    IConfigurationExtend,
    IConfigurationModify,
    IEnvironmentRead,
    IHttp,
    ILogger,
    IMessageBuilder,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    ApiSecurity,
    ApiVisibility,
} from "@rocket.chat/apps-engine/definition/api";
import { App } from "@rocket.chat/apps-engine/definition/App";
import {
    IMessage,
    IMessageDeleteContext,
    IPostMessageDeleted,
    IPostMessageSent,
    IPostMessageUpdated,
    IPreMessageDeletePrevent,
    IPreMessageSentModify,
    IPreMessageSentPrevent,
    IPreMessageUpdatedPrevent,
} from "@rocket.chat/apps-engine/definition/messages";
import { IAppInfo } from "@rocket.chat/apps-engine/definition/metadata";
import {
    IPostRoomUserJoined,
    IPreRoomUserLeave,
    IRoom,
    IRoomUserJoinedContext,
    IRoomUserLeaveContext,
} from "@rocket.chat/apps-engine/definition/rooms";
import {
    IJobContext,
    StartupType,
} from "@rocket.chat/apps-engine/definition/scheduler";
import { ISetting } from "@rocket.chat/apps-engine/definition/settings";
import {
    RoomTypeFilter,
    UIActionButtonContext,
} from "@rocket.chat/apps-engine/definition/ui";
import {
    IUIKitResponse,
    UIKitActionButtonInteractionContext,
    UIKitBlockInteractionContext,
    UIKitViewSubmitInteractionContext,
} from "@rocket.chat/apps-engine/definition/uikit";
import {
    IFileUploadContext,
    IPreFileUpload,
} from "@rocket.chat/apps-engine/definition/uploads";
import { UserType } from "@rocket.chat/apps-engine/definition/users";
import { settings } from "./config/Settings";
import { AuthenticationEndpoint } from "./endpoints/AuthenticationEndpoint";
import { SubscriberEndpoint } from "./endpoints/SubscriberEndpoint";
import {
    IncomingNotificationProcessorId,
    OAuthNonceCleanupInterval,
    OAuthNonceCleanupJobId,
    RegistrationAutoRenewInterval,
    RegistrationAutoRenewSchedulerId,
    SubscriberEndpointPath,
    UIActionId,
    UIElementId,
    WebhookSecretCreationJobId,
} from "./lib/Const";
import {
    handleAddTeamsUserContextualBarSubmitAsync,
    handlePostMessageDeletedAsync,
    handlePostMessageSentAsync,
    handlePostMessageUpdatedAsync,
    handlePostRoomUserJoinedAsync,
    handlePreFileUploadAsync,
    handlePreMessageOperationPreventAsync,
    handlePreMessageSentPreventAsync,
    handlePreRoomUserLeaveAsync,
    handleUninstallApp,
    handleUserRegistrationAutoRenewAsync,
} from "./lib/EventHandler";
import { getRocketChatAppEndpointUrl } from "./lib/UrlHelper";
import {
    decodeButtonState,
    getRoomIdFromSubmitActionId,
    openAddTeamsUserContextualBarBlocksAsync,
    openViewTeamsMembersContextualBarAsync,
    updateAddTeamsUserContextualBarAsync,
    updateViewTeamsMembersContextualBarAsync,
} from "./lib/UserInterfaceHelper";
import { AddUserSlashCommand } from "./slashcommands/AddUserSlashCommand";
import { DeleteTeamsBotUserSlashCommand } from "./slashcommands/DeleteTeamsBotUserSlashCommand";
import { LoginTeamsSlashCommand } from "./slashcommands/LoginTeamsSlashCommand";
import { LogoutTeamsSlashCommand } from "./slashcommands/LogoutTeamsSlashCommand";
import { SetupVerificationSlashCommand } from "./slashcommands/SetupVerificationSlashCommand";
import { LoginAppUserSlashCommand } from "./slashcommands/LoginAppUserSlashCommand";
import { ResubscribeMessages } from "./slashcommands/ResubscriptionMessages";
import { ViewTeamsMembersSlashCommand } from "./slashcommands/ViewTeamsMembersSlashCommand";
import { OAuthNonce, SubscriptionRenewalJob, WebhookSecret } from "./lib/PersistHelper";
import { PreventRegistry } from "./lib/PreventRegistry";
import {
    getExtraInfoAndOriginalFileName,
    popExtraInfoAttachment,
} from "./lib/MessageHelper";
import { handleInboundNotificationAsync } from "./lib/InboundNotificationHelper";
import { handleWebhookSecretCreationAsync } from "./lib/handlers/handleWebhookSecretCreation";

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
        IPreMessageSentModify,
        IPreRoomUserLeave,
        IPostRoomUserJoined
{
    constructor(info: IAppInfo, logger: ILogger, accessors: IAppAccessors) {
        super(info, logger, accessors);
    }

    async getSettingValueById(id: string) {
        return this.getAccessors()
            .environmentReader.getSettings()
            .getValueById(id);
    }

    async onInstall(
        context: IAppInstallationContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify,
    ): Promise<void> {
        await WebhookSecret.create({ persistence });
    }

    async executePreMessageSentModify(
        message: IMessage,
        builder: IMessageBuilder,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
    ) {
        let extraInfoData = popExtraInfoAttachment(message);
        if (extraInfoData.source === "ms-teams") {
            await PreventRegistry.set(
                persistence,
                `PreventPostMessageHook/${message.id}`,
                true,
            );
            return message;
        }

        extraInfoData = {};
        let pos = -1;
        let originalFilenameFound = "";
        message.attachments?.forEach((att, i) => {
            if (!att.title?.value) {
                return false;
            }
            const { originalFilename, present, extraInfo } =
                getExtraInfoAndOriginalFileName(att.title.value);
            if (present) {
                extraInfoData = extraInfo;
                pos = i;
                originalFilenameFound = originalFilename;
            }
            return true;
        });

        if (extraInfoData.source === "ms-teams") {
            await PreventRegistry.set(
                persistence,
                `PreventPostMessageHook/${message.id}`,
                true,
            );
        }

        if (pos !== -1) {
            if (Array.isArray(message["_unmappedProperties_"]?.["files"])) {
                message["_unmappedProperties_"]["files"] = message[
                    "_unmappedProperties_"
                ]["files"].map((file) => {
                    if (file.name === message.attachments![pos].title?.value) {
                        return { ...file, name: originalFilenameFound };
                    }
                    return file;
                });
            }

            // Replace the attachment at position 'pos' with updated title
            if (message.attachments?.[pos]) {
                message.attachments[pos] = {
                    ...message.attachments[pos],
                    title: {
                        ...message.attachments[pos].title,
                        value: originalFilenameFound,
                    },
                };
            }

            // Update the file name if present
            if (message.file && originalFilenameFound) {
                message.file = { ...message.file, name: originalFilenameFound };
            }
            return message;
        }
        return message;
    }

    async onEnable(
        environment: IEnvironmentRead,
        configurationModify: IConfigurationModify,
    ): Promise<boolean> {
        try {
            await configurationModify.scheduler.scheduleOnce({
                id: WebhookSecretCreationJobId,
                when: new Date(),
                data: { from: "ScheduleOnce/Immediate" },
            });

            await configurationModify.scheduler.scheduleOnce({
                id: RegistrationAutoRenewSchedulerId,
                when: new Date(Date.now() + 5000),
                data: { from: "ScheduleOnce/5seconds" },
            });

            await configurationModify.scheduler.scheduleRecurring({
                id: RegistrationAutoRenewSchedulerId,
                interval: RegistrationAutoRenewInterval,
                data: { from: "ScheduleRecurring" },
                skipImmediate: true,
            });

            await configurationModify.scheduler.scheduleRecurring({
                id: OAuthNonceCleanupJobId,
                interval: OAuthNonceCleanupInterval,
                skipImmediate: true,
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
        return handleUninstallApp({
            read,
            http,
            modify,
            app: this,
            persistence,
        });
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
            modify,
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
            http,
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

    public async executePostRoomUserJoined(
        context: IRoomUserJoinedContext,
        read: IRead,
        http: IHttp,
        persistence: IPersistence,
        modify: IModify,
    ): Promise<void> {
        await handlePostRoomUserJoinedAsync({
            app: this,
            context,
            read,
            persistence,
            http,
            modify,
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
                throw new Error("App user not found");
            }

            await openAddTeamsUserContextualBarBlocksAsync(
                data.triggerId,
                data.room,
                data.user,
                appUser,
                read,
                modify,
                http,
                persistence,
                this,
            );
        }

        if (data.actionId === UIActionId.ViewTeamsMembersButtonClicked) {
            await openViewTeamsMembersContextualBarAsync(
                data.triggerId,
                data.room,
                data.user,
                read,
                modify,
                http,
                persistence,
                this,
            );
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
        modify: IModify,
    ): Promise<IUIKitResponse> {
        const { actionId, value, room } = context.getInteractionData();

        if (actionId === UIActionId.TeamsUserSearchInput) {
            // Fires on every keystroke (ON_CHARACTER_ENTERED dispatch).
            // value = current text in the search input; room = the open room.
            const roomId = room?.id ?? '';
            const updatedView = await updateAddTeamsUserContextualBarAsync({
                actionId,
                value,
                roomId,
                read,
                http,
                persistence,
                app: this,
                modify,
            });
            if (updatedView) {
                return context.getInteractionResponder().updateContextualBarViewResponse(updatedView);
            }
        }

        if (actionId === UIActionId.TeamsUserLoadMore && value) {
            // value = base64 encoded { nextLink, loadedUsers, roomId }.
            const { roomId } = decodeButtonState(value);
            const updatedView = await updateAddTeamsUserContextualBarAsync({
                actionId,
                value,
                roomId,
                read,
                http,
                persistence,
                app: this,
                modify,
            });
            if (updatedView) {
                return context.getInteractionResponder().updateContextualBarViewResponse(updatedView);
            }
        }

        if (actionId === UIActionId.ViewMembersLoadMore && value) {
            // value = base64 encoded { nextLink, loadedMembers, threadId, total }.
            const updatedView = await updateViewTeamsMembersContextualBarAsync({
                value,
                read,
                http,
                persistence,
                app: this,
                modify,
            });
            if (updatedView) {
                return context.getInteractionResponder().updateContextualBarViewResponse(updatedView);
            }
        }

        return context.getInteractionResponder().successResponse();
    }

    public async executeViewClosedHandler(): Promise<IUIKitResponse> {
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
            const roomIdFromActionId =
                submitActionId && getRoomIdFromSubmitActionId(submitActionId);
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
                    });
                });
            }

            // Fallback to object property implementation
            if (teamsUserIdsToSave && currentRoom) {
                await handleAddTeamsUserContextualBarSubmitAsync({
                    operator: user,
                    room: currentRoom,
                    teamsUserIdsToSave,
                    read,
                    persistence,
                    http,
                    modify,
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
            modify.getDeleter().deleteUsers(this.getID(), UserType.BOT), // To remove old bot users
        ]);
        return;
    }

    protected incomingNotificationJob = async (
        jobContext: IJobContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persistence: IPersistence,
    ): Promise<void> => {
        await handleInboundNotificationAsync({
            app: this,
            read,
            modify,
            http,
            inBoundNotification: jobContext.inBoundNotification,
            persistence,
        });
    };

    protected webhookSecretCreationJob = async (
        jobContext: IJobContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persistence: IPersistence,
    ): Promise<void> => {
        handleWebhookSecretCreationAsync({
            app: this,
            read,
            http,
            persistence,
        });
    };

    protected registrationRenewalsJob = async (
        jobContext: IJobContext,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persistence: IPersistence,
    ) => {
        try {
            console.log(
                `[Teams Bridge] Start renew registrations! (from: ${jobContext.from})`,
            );
            let jobState = await SubscriptionRenewalJob.find({
                persistenceRead: read.getPersistenceReader(),
            });

            if (
                jobState &&
                jobState.lastStartedJobTimestamp &&
                Date.now() -
                    new Date(
                        jobState.lastStartedJobTimestamp,
                    ).getTime() <
                    5 * 60 * 1000
            ) {
                // Job ran less than 5 minutes ago
                console.log(
                    `[Teams Bridge] ${RegistrationAutoRenewSchedulerId} Job already ran less than 5 minutes ago. Skipping this run.`,
                );
                return;
            }

            await SubscriptionRenewalJob.persist({
                persistence,
                lastStartedJobTimestamp: new Date(),
            });

            const subscriberEndpointUrl =
                await getRocketChatAppEndpointUrl(
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
            console.log(
                "[Teams Bridge] Finish renew registrations!",
            );
        } catch (error) {
            throw new Error(
                `[Teams Bridge] Auto renew registration failed with error: ${error}`,
            );
        }
    };

    protected oauthNonceCleanupJob = async (
        _jobContext: IJobContext,
        read: IRead,
        _modify: IModify,
        _http: IHttp,
        persistence: IPersistence,
    ) => {
        try {
            await OAuthNonce.deleteStale(read, persistence);
        } catch (error) {
            throw new Error(
                `[Teams Bridge] OAuth nonce cleanup failed with error: ${error}`,
            );
        }
    };

    protected async extendConfiguration(
        configuration: IConfigurationExtend,
    ): Promise<void> {
        // Register app settings
        await Promise.all(
            settings.map((setting) =>
                configuration.settings.provideSetting(setting),
            ),
        );

        await Promise.all([
            configuration.slashCommands.provideSlashCommand(
                new SetupVerificationSlashCommand(),
            ),
            configuration.slashCommands.provideSlashCommand(
                new DeleteTeamsBotUserSlashCommand(this),
            ),
            configuration.slashCommands.provideSlashCommand(
                new LoginTeamsSlashCommand(this),
            ),
            configuration.slashCommands.provideSlashCommand(
                new LogoutTeamsSlashCommand(this),
            ),
            configuration.slashCommands.provideSlashCommand(
                new AddUserSlashCommand(this),
            ),
            configuration.slashCommands.provideSlashCommand(
                new ResubscribeMessages(this),
            ),
            configuration.slashCommands.provideSlashCommand(
                new LoginAppUserSlashCommand(this),
            ),
            configuration.slashCommands.provideSlashCommand(
                new ViewTeamsMembersSlashCommand(this),
            ),
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

        // Config context menu item - Add Teams user
        configuration.ui.registerButton({
            actionId: UIActionId.AddTeamsUserButtonClicked,
            labelI18n: "action_button_label_add_teams_user",
            context: UIActionButtonContext.ROOM_ACTION,
            when: {
                roomTypes: [
                    RoomTypeFilter.PRIVATE_DISCUSSION,
                    RoomTypeFilter.PRIVATE_CHANNEL,
                    RoomTypeFilter.PRIVATE_TEAM,
                ],
            },
        });

        // Config context menu item - View Teams members
        configuration.ui.registerButton({
            actionId: UIActionId.ViewTeamsMembersButtonClicked,
            labelI18n: "action_button_label_view_teams_members",
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
                processor: this.registrationRenewalsJob,
            },
            {
                id: WebhookSecretCreationJobId,
                processor: this.webhookSecretCreationJob,
            },
            {
                id: IncomingNotificationProcessorId,
                processor: this.incomingNotificationJob,
            },
            {
                id: OAuthNonceCleanupJobId,
                processor: this.oauthNonceCleanupJob,
            },
        ]);
    }
}
