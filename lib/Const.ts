export const AuthenticationEndpointPath: string = 'auth';

export const MicrosoftBaseUrl: string = 'https://login.microsoftonline.com';

export const LoginMessageText: string =
    'To start cross platform collaboration, you need to login to Microsoft with your Teams account or guest account. '
    + 'You\'ll be able to keep using Rocket.Chat, but you\'ll also be able to chat with colleagues using Microsoft Teams. '
    + 'Please click this button to login Teams:';

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
    'chatmessage.send'
];

export const getMicrosoftTokenUrl = (aadTenantId: string) => {
    return `${MicrosoftBaseUrl}/${aadTenantId}/oauth2/v2.0/token`;
};

export const getMicrosoftAuthorizeUrl = (aadTenantId: string) => {
    return `${MicrosoftBaseUrl}/${aadTenantId}/oauth2/v2.0/authorize`;
};
