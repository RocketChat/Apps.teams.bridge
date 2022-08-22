import { UserModel } from "./PersistHelper";

export const AuthenticationEndpointPath: string = 'auth';

const LoginBaseUrl: string = 'https://login.microsoftonline.com';
const GraphApiBaseUrl: string = 'https://graph.microsoft.com';
export const SupportDocumentUrl: string = 'https://github.com/RocketChat/Apps.teams.bridge/blob/main/docs/support.md';

const GraphApiVersion = {
    V1: 'v1.0',
    Beta: 'beta',
}

const GraphApiEndpoint = {
    Profile: 'me',
    Chat: 'chats',
    Message: (threadId: string) => `chats/${threadId}/messages`,
};

export const LoginMessageText: string =
    'To start cross platform collaboration, you need to login to Microsoft with your Teams account or guest account. '
    + 'You\'ll be able to keep using Rocket.Chat, but you\'ll also be able to chat with colleagues using Microsoft Teams. '
    + 'Please click this button to login Teams:';
export const LoginRequiredHintMessageText: string = 
    'The Rocket.Chat user you are messaging represents a colleague in your organization using Microsoft Teams. '
    + 'The message can NOT be delivered to the user on Microsoft Teams before you start cross platform collaboration for your account. '
    + 'For details, see:';

export const LoginButtonText: string = 'Login Teams';
export const SupportDocumentButtonText: string = 'Support Document';

export const AuthenticationScopes = [
    'offline_access',
    'user.read.all',
    'chat.create',
    'chat.readbasic',
    'chat.readwrite',
    'chatmember.read',
    'chatmember.readwrite',
    'chatmessage.read',
    'chatmessage.send'
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

export const getGraphApiMessageUrl = (threadId: string) => {
    return `${GraphApiBaseUrl}/${GraphApiVersion.V1}/${GraphApiEndpoint.Message(threadId)}`;
};

export const TestEnvironment = {
    // Set enable to true for local testing with mock data
    enable: true,
    // Put url here when running locally & using tunnel service such as Ngrok to expose the localhost port to the internet
    tunnelServiceUrl: '',
    mockDummyUsers: [
        {
            // Mock dummy user for alexw.l4cf.onmicrosoft.com 
            rocketChatUserId: 'v4ECCH3pTAE6nBXyJ',
            teamsUserId: 'ffa3322f-670c-4887-b193-a04cca6073f8',
        }
    ] as UserModel[],
};
