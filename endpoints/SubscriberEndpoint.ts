import {
    IHttp,
    IModify,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    ApiEndpoint,
    IApiEndpointInfo,
    IApiRequest,
    IApiResponse,
} from "@rocket.chat/apps-engine/definition/api";
import { IApp } from "@rocket.chat/apps-engine/definition/IApp";
import { SubscriberEndpointPath } from "../lib/Const";
import {
    handleInboundNotificationAsync,
    InBoundNotification,
    NotificationChangeType,
    NotificationResourceType,
} from "../lib/InboundNotificationHelper";

interface RawNotification {
    changeType: string;
    resourceData: {
        "@odata.type": string;
        id: string;
    };
    resource: string;
    subscriptionId: string;
}

export class SubscriberEndpoint extends ApiEndpoint {
    private supportedChangeTypeMapping: Record<string, NotificationChangeType> = {
        created: NotificationChangeType.Created,
        updated: NotificationChangeType.Updated,
        deleted: NotificationChangeType.Deleted,
    };

    private supportedResourceTypeMapping: Record<string, NotificationResourceType> = {
        "#Microsoft.Graph.chatMessage": NotificationResourceType.ChatMessage,
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
        persis: IPersistence
    ): Promise<IApiResponse> {
        if (request.query?.validationToken) {
            return this.success(request.query.validationToken);
        }

        const receiverRocketChatUserId: string = request.query.userId;
        const notifications: RawNotification[] = request.content.value || [];

        for (const rawNotification of notifications) {
            try {
                const changeType = this.parseChangeType(rawNotification.changeType);
                const resourceType = this.parseResourceType(rawNotification.resourceData["@odata.type"]);

                if (!changeType || !resourceType) {
                    continue;
                }

                const inBoundNotification: InBoundNotification = {
                    receiverRocketChatUserId,
                    subscriptionId: rawNotification.subscriptionId,
                    changeType,
                    resourceId: rawNotification.resourceData.id,
                    resourceString: rawNotification.resource,
                    resourceType,
                };

                await handleInboundNotificationAsync(
                    inBoundNotification,
                    read,
                    modify,
                    http,
                    persis,
                    this.app.getID()
                );
            } catch (error) {
                console.error(`Error when handling inbound notification. Details: ${error}`);
            }
        }

        return this.success("OK");
    }

    private parseChangeType(changeType: string): NotificationChangeType | undefined {
        return this.supportedChangeTypeMapping[changeType];
    }

    private parseResourceType(resourceType: string): NotificationResourceType | undefined {
        return this.supportedResourceTypeMapping[resourceType];
    }
}
