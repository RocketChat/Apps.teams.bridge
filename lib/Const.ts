import { UserModel } from "./PersistHelper";

export const AuthenticationEndpointPath: string = 'auth';
export const SubscriberEndpointPath: string = 'subscriber';

const LoginBaseUrl: string = 'https://login.microsoftonline.com';
const GraphApiBaseUrl: string = 'https://graph.microsoft.com';

export const SubscriptionMaxExpireTimeInSecond: number = 3600;

const GraphApiVersion = {
    V1: 'v1.0',
    Beta: 'beta',
};

const GraphApiEndpoint = {
    Profile: 'me',
    Chat: 'chats',
    Message: (threadId: string) => `chats/${threadId}/messages`,
    DeleteMessage: (userId: string, threadId: string, messageId: string) => 
        `users/${userId}/chats/${threadId}/messages/${messageId}/softDelete`,
    Subscription: 'subscriptions',
};

export const AppSetupVerificationPassMessageText: string = 'TeamsBridge app setup verification PASSED!';
export const AppSetupVerificationFailMessageText: string =
    'TeamsBridge app setup verification FAILED! Please check trouble shooting guide for further actions.';
export const LoginMessageText: string =
    'To start cross platform collaboration, you need to login to Microsoft with your Teams account or guest account. '
    + 'You\'ll be able to keep using Rocket.Chat, but you\'ll also be able to chat with colleagues using Microsoft Teams. '
    + 'Please click this button to login Teams:';
export const LoginRequiredHintMessageText: string = 
    'The Rocket.Chat user you are messaging represents a colleague in your organization using Microsoft Teams. '
    + 'The message can NOT be delivered to the user on Microsoft Teams before you start cross platform collaboration for your account. '
    + 'Please click this button to login Teams:';
export const LoggedInBridgeUserRequiredHintMessageText: string = 
    'The Rocket.Chat room you are messaging includes at least one dummy user that represents a colleague in your organization using Microsoft Teams. '
    + 'The message can NOT be delivered to Microsoft Teams before there is at least one user in this room start cross platform collaboration. '
    + 'To start cross platform collaboration for your account, please click this button to login Teams:';
export const UnsupportedScenarioHintMessageText = (scenario: string) =>
    `${scenario} is not supported by TeamsBridge app for cross platform collaboration.`
    + ' This message won\'t be delivered to target user on Teams.';
export const BridgeUserNotificationMessageText: string = 
    'This Rocket.Chat room includes at least one dummy user that represents a colleague in your organization using Microsoft Teams. '
    + 'You have became the bridge user of this room. '
    + 'All messages sent by unlogged-in user to this room will be delivered to Microsoft Teams by you.';

export const LoginButtonText: string = 'Login Teams';

export const AuthenticationScopes = [
    'offline_access',
    'user.read.all',
    'chat.create',
    'chat.readbasic',
    'chat.readwrite',
    'chatmember.read',
    'chatmember.readwrite',
    'chatmessage.read',
    'chatmessage.send',
    'files.readwrite',
];

export const SupportedNotificationChangeTypes = [
    'created',
    'updated',
    // 'deleted',
];

export const getMicrosoftTokenUrl = (aadTenantId: string) => {
    return `${LoginBaseUrl}/${aadTenantId}/oauth2/v2.0/token`;
};

export const getMicrosoftAuthorizeUrl = (aadTenantId: string) => {
    return `${LoginBaseUrl}/${aadTenantId}/oauth2/v2.0/authorize`;
};

export const getGraphApiProfileUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Profile}`;
};

export const getGraphApiChatUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Chat}`;
};

export const getGraphApiMessageUrl = (
    threadId: string,
    messageId?: string,
    useBetaVersion?: boolean) => {
    let version = GraphApiVersion.V1;
    if (useBetaVersion) {
        version = GraphApiVersion.Beta;
    }

    let url = `${GraphApiBaseUrl}/${version}/${GraphApiEndpoint.Message(threadId)}`;
    if (messageId) {
        url = `${url}/${messageId}`;
    }

    return url;
};

export const getGraphApiMessageDeleteUrl = (
    userId: string,
    threadId: string,
    messageId: string) => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.Beta}/${GraphApiEndpoint.DeleteMessage(userId, threadId, messageId)}`;
};

export const getGraphApiMessageBetaUrl = (threadId: string, messageId?: string) => {
    if (messageId) {
        return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Message(threadId)}/${messageId}`;
    }

    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Message(threadId)}`;
};

export const getGraphApiSubscriptionUrl = () => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Subscription}`;
};

export const getGraphApiResourceUrl = (resourceString: string) => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${resourceString}`;
};

export const TestEnvironment = {
    // Set enable to true for local testing with mock data
    enable: true,
    // Put url here when running locally & using tunnel service such as Ngrok to expose the localhost port to the internet
    tunnelServiceUrl: 'https://cce3-50-35-80-12.ngrok.io',
    mockDummyUsers: [
        {
            // Mock dummy user for alexw.l4cf.onmicrosoft.com 
            rocketChatUserId: 'LXrarr3m2NJi9fSQd',
            teamsUserId: 'ffa3322f-670c-4887-b193-a04cca6073f8',
        }
    ] as UserModel[],
};
