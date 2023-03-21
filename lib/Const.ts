export const AuthenticationEndpointPath: string = "auth";
export const SubscriberEndpointPath: string = "subscriber";

const LoginBaseUrl: string = "https://login.microsoftonline.com";
const GraphApiBaseUrl: string = "https://graph.microsoft.com";

export const SubscriptionMaxExpireTimeInSecond: number = 3600;

const GraphApiVersion = {
    V1: "v1.0",
    Beta: "beta",
};

const GraphApiEndpoint = {
    Profile: "me",
    RevokeRefreshToken: "me/revokeSignInSessions",
    User: "users",
    Chat: "chats",
    ChatThread: (threadId: string) => `chats/${threadId}`,
    ChatMember: (threadId: string) => `chats/${threadId}/members`,
    RemoveChatMember: (threadId: string, userId: string) =>
        `chats/${threadId}/members/${userId}`,
    Message: (threadId: string) => `chats/${threadId}/messages`,
    DeleteMessage: (userId: string, threadId: string, messageId: string) =>
        `users/${userId}/chats/${threadId}/messages/${messageId}/softDelete`,
    Subscription: "subscriptions",
    SubscriptionOperation: (subscriptionId: string) =>
        `subscriptions/${subscriptionId}`,
    Share: (encodedUrl: string) => `shares/${encodedUrl}/driveItem/content`,
    Upload: (filename: string) =>
        `/me/drive/items/root:/rocketchatshare/${filename}:/content`,
    OneDriveItem: (driveItemId: string) => `/me/drive/items/${driveItemId}`,
    ShareLink: (driveItemId: string) =>
        `/me/drive/items/${driveItemId}/createLink`,
};

export const AppSetupVerificationPassMessageText: string =
    "TeamsBridge app setup verification PASSED!";
export const AppSetupVerificationFailMessageText: string =
    "TeamsBridge app setup verification FAILED! Please check trouble shooting guide for further actions.";
export const ProvisionTeamsBotUserSucceedMessageText: string =
    "Provision Teams bot user succeed!";
export const ProvisionTeamsBotUserFailedMessageText: string =
    "Provision Teams bot user FAILED! Please check trouble shooting guide for further actions.";
export const LoginMessageText: string =
    "To start cross platform collaboration, you need to login to Microsoft with your Teams account or guest account. " +
    "You'll be able to keep using Rocket.Chat, but you'll also be able to chat with colleagues using Microsoft Teams. " +
    "Please click this button to login Teams:";
export const LoginRequiredHintMessageText: string =
    "The Rocket.Chat user you are messaging represents a colleague in your organization using Microsoft Teams. " +
    "The message can NOT be delivered to the user on Microsoft Teams before you start cross platform collaboration for your account. " +
    "Please click this button to login Teams:";
export const LoginNoNeedHintMessageText: string =
    "You have already login Microsoft Teams to start cross platform collaboration for your account. " +
    "No need to login again.";
export const LoggedInBridgeUserRequiredHintMessageText: string =
    "The Rocket.Chat room you are messaging includes at least one Teams Bot user that represents a colleague in your organization using Microsoft Teams. " +
    "The message can NOT be delivered to Microsoft Teams before there is at least one user in this room start cross platform collaboration. " +
    "To start cross platform collaboration for your account, please click this button to login Teams:";
export const UnsupportedScenarioHintMessageText = (scenario: string) =>
    `${scenario} is not supported by TeamsBridge app for cross platform collaboration.` +
    " This message won't be delivered to target user on Teams.";
export const BridgeUserNotificationMessageText: string =
    "This Rocket.Chat room includes at least one Teams Bot user that represents a colleague in your organization using Microsoft Teams. " +
    "You have became the bridge user of this room. " +
    "All messages sent by unlogged-in user to this room will be delivered to Microsoft Teams by you.";
export const AddUserRoomTypeInvalidHintMessageText: string =
    "Adding a Teams Bot user only supported for private channels, private teams, and private discussions.";
export const AddUserNoExistingUsersHintMessageText: string =
    "No Teams Bot user provisioned for your organization. Please contact your organization admin for help.";
export const AddUserNameInvalidHintMessageText: string =
    "The user you are trying to add does not exist or is not a Teams Bot user. Please find the correct user name a try again.";
export const AddUserLoginRequiredHintMessageText: string =
    "Adding a Teams Bot user to a group chat room requires at least one user in this room start cross platform collaboration. " +
    "To start cross platform collaboration for your account, please click this button to login Teams:";
export const LogoutNoNeedHintMessageText: string =
    "You have NOT logged in to Microsoft yet. No need to logout.";
export const LogoutSuccessHintMessageText: string =
    "You have successfully logged out Microsoft Teams.";

export const LoginButtonText: string = "Login Teams";

export const AuthenticationScopes = [
    "offline_access",
    "user.read.all",
    "chat.create",
    "chat.readbasic",
    "chat.readwrite",
    "chatmember.read",
    "chatmember.readwrite",
    "chatmessage.read",
    "chatmessage.send",
    "files.readwrite",
];

export const SupportedNotificationChangeTypes = [
    "created",
    "updated",
    "deleted",
];

export const MicrosoftFileUrlPrefix = "https://graph.microsoft.com";

export const SharePointUrl = "sharepoint.com";

export const FileAttachmentContentType = "reference";

export const TeamsAppUserNameSurfix = "msteams.alias";

export const RegistrationAutoRenewSchedulerId =
    "registration.auto.renew.scheduler";

export const RegistrationAutoRenewInterval = "1800 seconds";

export const DefaultThreadName = "Rocket.Chat interop group";

export const DefaultTeamName = "TeamsInteropGroupChat";

export const UIActionId = {
    AddTeamsUserButtonClicked: "TeamsBridge.AddTeamsUserButtonClicked",
    TeamsUserNameSearch: "TeamsBridge.TeamsUserNameSearch",
    SaveChanges: "TeamsBridge.SaveChanges",
};

export const UIElementId = {
    ContextualBarId: "TeamsBridge.ContextualBarId",
};

export const UIElementText = {
    ContextualBarTitle: "Add Teams users",
    TeamsUserNameSearchTitle: "Choose Teams users",
    TeamsUserNameSearchPlaceHolder: "Choose Teams users",
    TeamsUsersSaveChangeButton: "Add users",
};

export const getMicrosoftTokenUrl = (aadTenantId: string) => {
    return `${LoginBaseUrl}/${aadTenantId}/oauth2/v2.0/token`;
};

export const getMicrosoftAuthorizeUrl = (aadTenantId: string) => {
    return `${LoginBaseUrl}/${aadTenantId}/oauth2/v2.0/authorize`;
};

export const getGraphApiRevokeRefreshTokenUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.RevokeRefreshToken}`;
};

export const getGraphApiProfileUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Profile}`;
};

export const getGraphApiUserUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.User}`;
};

export const getGraphApiChatUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Chat}`;
};

export const getGraphApiChatThreadWithMemberUrl = (threadId: string) => {
    return `${GraphApiBaseUrl}/${
        GraphApiVersion.V1
    }/${GraphApiEndpoint.ChatThread(threadId)}?$expand=members`;
};

export const getGraphApiChatMemberUrl = (threadId: string) => {
    return `${GraphApiBaseUrl}/${
        GraphApiVersion.V1
    }/${GraphApiEndpoint.ChatMember(threadId)}`;
};

export const getGraphApiChatMemberRemoveUrl = (
    threadId: string,
    userId: string
) => {
    return `${GraphApiBaseUrl}/${
        GraphApiVersion.V1
    }/${GraphApiEndpoint.RemoveChatMember(threadId, userId)}`;
};

export const getGraphApiMessageUrl = (
    threadId: string,
    messageId?: string,
    useBetaVersion?: boolean
) => {
    let version = GraphApiVersion.V1;
    if (useBetaVersion) {
        version = GraphApiVersion.Beta;
    }

    let url = `${GraphApiBaseUrl}/${version}/${GraphApiEndpoint.Message(
        threadId
    )}`;
    if (messageId) {
        url = `${url}/${messageId}`;
    }

    return url;
};

export const getGraphApiShareUrl = (encodedUrl: string) => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Share(
        encodedUrl
    )}`;
};

export const getGraphApiUploadToDriveUrl = (fileName: string) => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Upload(
        fileName
    )}`;
};

export const getGraphApiShareOneDriveFileUrl = (oneDriveItemId: string) => {
    return `${GraphApiBaseUrl}/${
        GraphApiVersion.V1
    }/${GraphApiEndpoint.ShareLink(oneDriveItemId)}`;
};

export const getGraphApiOneDriveFileLinkUrl = (oneDriveItemId: string) => {
    return `${GraphApiBaseUrl}/${
        GraphApiVersion.V1
    }/${GraphApiEndpoint.OneDriveItem(oneDriveItemId)}`;
};

export const getGraphApiMessageDeleteUrl = (
    userId: string,
    threadId: string,
    messageId: string
) => {
    return `${GraphApiBaseUrl}/${
        GraphApiVersion.Beta
    }/${GraphApiEndpoint.DeleteMessage(userId, threadId, messageId)}`;
};

export const getGraphApiMessageBetaUrl = (
    threadId: string,
    messageId?: string
) => {
    if (messageId) {
        return `${GraphApiBaseUrl}/${
            GraphApiVersion.V1
        }/${GraphApiEndpoint.Message(threadId)}/${messageId}`;
    }

    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Message(
        threadId
    )}`;
};

export const getGraphApiSubscriptionUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Subscription}`;
};

export const getGraphApiSubscriptionOperationUrl = (subscriptionId: string) => {
    return `${GraphApiBaseUrl}/${
        GraphApiVersion.V1
    }/${GraphApiEndpoint.SubscriptionOperation(subscriptionId)}`;
};

export const getGraphApiResourceUrl = (resourceString: string) => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${resourceString}`;
};

export const TestEnvironment = {
    // Put url here when running locally & using tunnel service such as Ngrok to expose the localhost port to the internet
    tunnelServiceUrl: "",
};
