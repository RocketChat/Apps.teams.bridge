import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IApiEndpointInfo, IApiRequest, IApiResponse } from '@rocket.chat/apps-engine/definition/api';
import { ApiEndpoint } from '@rocket.chat/apps-engine/definition/api';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { IncomingNotificationProcessorId, SubscriberEndpointPath } from '../lib/Const';
import { WebhookSecret } from '../lib/PersistHelper';
import type { InBoundNotification } from '../lib/inboundNotification/handleInboundNotificationAsync';
import { NotificationChangeType, NotificationResourceType } from '../lib/inboundNotification/handleInboundNotificationAsync';

export class SubscriberEndpoint extends ApiEndpoint {
	app: TeamsBridgeApp;

	private supportedChangeTypeMapping = {
		created: NotificationChangeType.Created,
		updated: NotificationChangeType.Updated,
		deleted: NotificationChangeType.Deleted,
	};

	private supportedResourceTypeMapping = {
		'#Microsoft.Graph.chatMessage': NotificationResourceType.ChatMessage,
	};

	public path = SubscriberEndpointPath;

	constructor(app: TeamsBridgeApp) {
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
		persis: IPersistence,
	): Promise<IApiResponse> {
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

				const clientState = rawNotification.clientState;

				if (!clientState) {
					// If clientState is not present, either it's an old subscription or
					// the notification is not from our app. We should ignore it.
					const message = `Source of notification cannot be verified. clientState is missing. Processing skipped.`;
					this.app.getLogger().error(message);
					return {
						status: 401,
						content: message,
					};
				}

				if (
					clientState !==
					(await WebhookSecret.getSubscriptionStateHash(read.getPersistenceReader(), persis, {
						rocketChatUserId: receiverRocketChatUserId,
					}))
				) {
					const message = `Source of notification cannot be verified. clientState is invalid. Processing skipped.`;
					this.app.getLogger().error(message);
					return {
						status: 401,
						content: message,
					};
				}

				const inBoundNotification: InBoundNotification = {
					receiverRocketChatUserId,
					subscriptionId: rawNotification.subscriptionId,
					changeType,
					resourceId: rawNotification.resourceData.id,
					resourceString: rawNotification.resource,
					resourceType,
				};

				await modify.getScheduler().scheduleOnce({
					when: new Date(),
					data: { inBoundNotification },
					id: IncomingNotificationProcessorId,
				});
			} catch (error) {
				// If there's an error, print a warning but not block the whole process
				console.error(`Error when handling inbound notification. Details: ${error.message}`);
			}
		}

		return this.success('OK');
	}

	private parseChangeType(changeType: string): NotificationChangeType | undefined {
		return this.supportedChangeTypeMapping[changeType];
	}

	private parseResourceType(resourceType: string): NotificationResourceType | undefined {
		return this.supportedResourceTypeMapping[resourceType];
	}
}
