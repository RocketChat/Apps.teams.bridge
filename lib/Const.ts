export const getMicrosoftOAuthUrl = (aadTenantId: string) => {
    return `https://login.microsoftonline.com/${aadTenantId}/oauth2/v2.0/token`;
};
