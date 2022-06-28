export const AuthenticationEndpointPath: string = 'auth';

export const MicrosoftBaseUrl: string = 'https://login.microsoftonline.com';

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
