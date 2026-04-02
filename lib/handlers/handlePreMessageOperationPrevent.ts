import type { IHttp, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import type { IMessage } from '@rocket.chat/apps-engine/definition/messages';

import type { TeamsBridgeApp } from '../../TeamsBridgeApp';

export const handlePreMessageOperationPreventAsync = async (_options: {
	message: IMessage;
	read: IRead;
	persistence: IPersistence;
	app: TeamsBridgeApp;
	http: IHttp;
}): Promise<boolean> => {
	// Single-bot architecture: no per-user Teams identity to check.
	// Edit/delete operations on bridged messages are controlled by the
	// message-ID mapping check in the update/delete handlers themselves.
	return false;
};
