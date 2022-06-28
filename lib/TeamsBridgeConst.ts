export const AuthenticationEndpointPath: string = 'auth';

export const MicrosoftLoginBaseUrl: string = 'https://login.microsoftonline.com';

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
