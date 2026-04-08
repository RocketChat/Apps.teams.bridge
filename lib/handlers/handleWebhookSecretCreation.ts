import type { IRead, IHttp, IPersistence } from '@rocket.chat/apps-engine/definition/accessors';

import type { TeamsBridgeApp } from '../../TeamsBridgeApp';
import { SubscriberEndpointPath } from '../Const';
import { getRocketChatAppEndpointUrl } from '../UrlHelper';
import { WebhookSecret } from '../persistence';
import { handleUserRegistrationAutoRenewAsync } from './handleUserRegistrationAutoRenew';

export const handleWebhookSecretCreationAsync = async ({
	app,
	read,
	http,
	persistence,
}: {
	app: TeamsBridgeApp;
	read: IRead;
	http: IHttp;
	persistence: IPersistence;
}): Promise<void> => {
	try {
		const webhookSecret = await WebhookSecret.get({
			persistenceRead: read.getPersistenceReader(),
		});
		if (!webhookSecret) {
			app.getLogger().info('Webhook secret is not created. Creating it now.');
			await WebhookSecret.create({ persistence });
			const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(app.getAccessors(), SubscriberEndpointPath);

			await handleUserRegistrationAutoRenewAsync({
				subscriberEndpointUrl,
				read,
				http,
				persistence,
				app,
			});
			app.getLogger().info('Webhook secret created and subscriptions were renewed.');
		}
	} catch (error) {
		app.getLogger().error(`Webhook secret creation failed with error, Incoming messages may fail to be processed`, error);
		throw new Error(`Webhook secret creation failed with error: ${error}`);
	}
};
