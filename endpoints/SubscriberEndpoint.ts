import {
    IHttp,
    IModify,
    IPersistence,
    IRead
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    ApiEndpoint,
    IApiEndpointInfo,
    IApiRequest,
    IApiResponse
} from "@rocket.chat/apps-engine/definition/api";
import { IApp } from "@rocket.chat/apps-engine/definition/IApp";
import { SubscriberEndpointPath } from "../lib/Const";
import {
    handleInboundNotificationAsync,
    InBoundNotification,
    NotificationChangeType,
    NotificationResourceType
} from "../lib/InboundNotificationHelper";

export class SubscriberEndpoint extends ApiEndpoint {
    private supportedChangeType = {
        'created': NotificationChangeType.Created,
    };

    private supportedResourceType = {
        '#Microsoft.Graph.chatMessage': NotificationResourceType.ChatMessage,
    };

    public path = SubscriberEndpointPath;

    constructor(app: IApp) {
        super(app);
        this.parseChangeType = this.parseChangeType.bind(this);
        this.parseResourceType = this.parseResourceType.bind(this);
    }

    public async post(
        request: IApiRequest,
        endpoint: IApiEndpointInfo,
        read: IRead,
        modify: IModify,
        http: IHttp,
        persis: IPersistence): Promise<IApiResponse> {
        console.log(request);
        if (request && request.query && request.query.validationToken) {
            return this.success(request.query.validationToken);
        }

        const receiverRocketChatUserId: string = request.query.userId;

        const notifications = request.content.value as any[];
        for (let index = 0; index < notifications.length; index++) {
            try {
                const rawNotification = notifications[index];
        
                const changeType = this.parseChangeType(rawNotification.changeType);
                if (!changeType) {
                    continue;
                }

                const resourceType = this.parseResourceType(rawNotification.resourceData['@odata.type']);
                if (!resourceType) {
                    continue;
                }

                const inBoundNotification : InBoundNotification = {
                    receiverRocketChatUserId: receiverRocketChatUserId,
                    subscriptionId: rawNotification.subscriptionId,
                    changeType: changeType,
                    resourceId: rawNotification.resourceData.id,
                    resourceString: rawNotification.resource,
                    resourceType: resourceType
                };

                await handleInboundNotificationAsync(
                    inBoundNotification,
                    read,
                    modify,
                    http,
                    persis);
            } catch (error) {
                // If there's an error, print a warning but not block the whole process
                console.log(`Error when handling inbound notification. Details: ${error}`);
            }
        }

        return this.success("OK");
    }

    private parseChangeType(changeType: string) : NotificationChangeType | undefined {
        return this.supportedChangeType[changeType];
    }

    private parseResourceType(resourceType: string) : NotificationResourceType | undefined {
        return this.supportedResourceType[resourceType];
    }
}