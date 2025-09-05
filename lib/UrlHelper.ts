import { IAppAccessors, IRead } from "@rocket.chat/apps-engine/definition/accessors";
import { IApiEndpointMetadata } from "@rocket.chat/apps-engine/definition/api";
import {
    AuthenticationScopes,
    getMicrosoftAuthorizeUrl,
    SubscriberEndpointPath,
} from "./Const";

import { AppSetting } from "../config/Settings";
import { IRoom } from "@rocket.chat/apps-engine/definition/rooms";

export const getRocketChatAppEndpointUrl = async (
    appAccessors: IAppAccessors,
    appEndpointPath: string
): Promise<string> => {
    const webhookEndpoint: IApiEndpointMetadata =
        appAccessors.providedApiEndpoints.find(
            (endpoint) => endpoint.path === appEndpointPath
        ) as IApiEndpointMetadata;
    let siteUrl: string = await appAccessors.environmentReader
        .getServerSettings()
        .getValueById("Site_Url");

    const proxyUrl = await appAccessors.environmentReader
        .getSettings()
        .getValueById(AppSetting.ProxyUrl);

    if (proxyUrl && proxyUrl !== "") {
        siteUrl = proxyUrl;
    }

    return new URL(webhookEndpoint.computedPath, siteUrl).toString();
};

export function getNotificationEndpointUrl(params: {
    appAccessors: IAppAccessors;
    rocketChatUserId: string;
}): Promise<string>;
export function getNotificationEndpointUrl(params: {
    rocketChatUserId: string;
    subscriberEndpoint: string;
}): string;
export function getNotificationEndpointUrl({
    appAccessors,
    rocketChatUserId,
    subscriberEndpoint,
}: {
    appAccessors?: IAppAccessors;
    rocketChatUserId?: string;
    subscriberEndpoint?: string;
}): Promise<string> | string {
    if (appAccessors) {
        return new Promise(async (resolve) => {
            const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(appAccessors, SubscriberEndpointPath);
            resolve(`${subscriberEndpointUrl}?userId=${rocketChatUserId}`);
        });
    } else if (subscriberEndpoint && rocketChatUserId) {
        return `${subscriberEndpoint}?userId=${rocketChatUserId}&hasClientState=1`;
    }
    throw new Error("Invalid parameters");
}


export const getLoginUrl = (
    aadTenantId: string,
    aadClientId: string,
    authEndpointUrl: string,
    userId: string
): string => {
    let url = getMicrosoftAuthorizeUrl(aadTenantId);
    url += `?client_id=${aadClientId}`;
    url += "&response_type=code";
    url += `&redirect_uri=${authEndpointUrl}`;
    url += "&response_mode=query";
    url += `&scope=${AuthenticationScopes.join("%20")}`;
    url += `&state=${userId}`;

    return url;
};


export const getRocketChatMessageUrl = async (read: IRead, msgId: string, room: IRoom) => {
    const siteUrl = await read.getEnvironmentReader().getServerSettings().getValueById("Site_Url");
    let roomType = 'channel';
    switch (room.type) {
        case 'p':
            roomType = 'group';
            break;
        case 'd':
            roomType = 'direct';
            break;
        case 'c':
        default:
            roomType = 'channel';
    }
    return `${siteUrl}/${roomType}/${room.id}?msg=${msgId}`;
}
