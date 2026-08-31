import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiChatUrl } from '../Const';

export interface MyChatMember {
	userId: string;
	displayName: string;
}

export interface MyChat {
	id: string; // Teams thread id (19:...@thread.v2)
	topic: string; // human label (chat topic, or members joined by comma for 1:1/group without topic)
	chatType: string;
	members: MyChatMember[];
}

export interface ListMyChatsResult {
	chats: MyChat[];
	nextLink?: string;
}

// Lists the signed-in (bot) account's Teams chats with members expanded, so an admin
// can pick an existing Teams conversation to link a Rocket.Chat room to.
export const listMyChatsAsync = async (http: IHttp, userAccessToken: string, options: { pageUrl?: string } = {}): Promise<ListMyChatsResult | null> => {
	const url = options.pageUrl ?? `${getGraphApiChatUrl()}?$expand=members&$top=50`;
	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(url, httpRequest);
	if (response.statusCode !== HttpStatusCode.OK) {
		return null;
	}

	const raw = (response.data?.value ?? []) as any[];
	const chats: MyChat[] = raw.map((c) => {
		const members: MyChatMember[] = ((c.members ?? []) as any[]).map((m) => ({
			userId: m.userId ?? '',
			displayName: m.displayName ?? m.email ?? m.userId ?? 'Unknown',
		}));
		let topic: string = c.topic ?? '';
		if (!topic) {
			// No explicit topic (1:1 or unnamed group): label with member names.
			topic = members.map((m) => m.displayName).filter(Boolean).join(', ') || c.id;
		}
		return { id: c.id, topic, chatType: c.chatType ?? '', members };
	});

	const nextLink: string | undefined = response.data?.['@odata.nextLink'] ?? undefined;
	return { chats, nextLink };
};
