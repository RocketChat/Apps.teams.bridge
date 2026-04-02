import type { App } from '@rocket.chat/apps-engine/definition/App';
import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IUserContext } from '@rocket.chat/apps-engine/definition/users';

import { LoginMessage, UserMapping, UserRegistration } from '../PersistHelper';

export const handlePostUserDeletedAsync = async ({
	app,
	context,
	read,
	persistence,
}: {
	app: App;
	context: IUserContext;
	read: IRead;
	persistence: IPersistence;
	http: IHttp;
	modify: IModify;
}): Promise<void> => {
	const { user } = context;
	if (!user?.id) return;

	try {
		await Promise.all([
			UserRegistration.delete(persistence, user.id),
			UserMapping.delete(read, persistence, user.id),
			LoginMessage.delete(persistence, user.id),
		]);
		app.getLogger().log(`[Teams Bridge] Successfully cleaned up Teams link mapping for deleted user ${user.id}.`);
	} catch (error) {
		app.getLogger().error(`[Teams Bridge] Error in handlePostUserDeletedAsync for user ${user.id}: ${(error as Error)?.message ?? String(error)}`);
	}
};
