import type { IHttp, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';

import type { TeamsBridgeApp } from '../TeamsBridgeApp';
import { AppSetting } from '../config/Settings';
import { UserMapping } from './PersistHelper';

interface BackupEntry {
	rcUsername: string;
	rcUserId: string;
	teamsUserId: string;
}

// Writes the full mapping table into the "Identity mappings backup" app setting as JSON.
// Keyed by RC username so a backup survives a full reinstall (user ids change, names don't).
export const syncMappingBackupAsync = async (options: { read: IRead; app: TeamsBridgeApp }): Promise<void> => {
	const { read, app } = options;
	try {
		const mappings = await UserMapping.findAll(read);
		const entries: BackupEntry[] = [];
		for (const m of mappings) {
			let rcUsername = '';
			try {
				const user = await read.getUserReader().getById(m.rocketChatUserId);
				rcUsername = user?.username ?? '';
			} catch (error) {
				// user gone — keep the id-only entry
			}
			entries.push({ rcUsername, rcUserId: m.rocketChatUserId, teamsUserId: m.teamsUserId });
		}
		entries.sort((a, b) => a.teamsUserId.localeCompare(b.teamsUserId));
		const json = JSON.stringify(entries, null, 2);

		const current = (await read.getEnvironmentReader().getSettings().getById(AppSetting.MappingsBackup)).value ?? '';
		if (current === json) {
			return;
		}
		await app.getAccessors().environmentWriter.getSettings().updateValue(AppSetting.MappingsBackup, json);
	} catch (error) {
		app.getLogger().warn(`Failed to sync mapping backup setting: ${error}`);
	}
};

// Imports mappings from the backup setting. Resolves each entry by RC username first
// (reinstall-safe), falling back to the stored RC user id. Returns counts for reporting.
export const restoreMappingBackupAsync = async (options: {
	read: IRead;
	persistence: IPersistence;
	http: IHttp;
	app: TeamsBridgeApp;
}): Promise<{ restored: number; skipped: number; error?: string }> => {
	const { read, persistence, app } = options;
	const raw = ((await read.getEnvironmentReader().getSettings().getById(AppSetting.MappingsBackup)).value ?? '').trim();
	if (!raw) {
		return { restored: 0, skipped: 0, error: 'The mappings backup setting is empty.' };
	}

	let entries: BackupEntry[];
	try {
		entries = JSON.parse(raw);
		if (!Array.isArray(entries)) {
			throw new Error('not an array');
		}
	} catch (error) {
		return { restored: 0, skipped: 0, error: 'The mappings backup setting does not contain valid JSON.' };
	}

	let restored = 0;
	let skipped = 0;
	for (const entry of entries) {
		const teamsUserId = entry?.teamsUserId;
		if (!teamsUserId) {
			skipped++;
			continue;
		}
		let rcUserId: string | undefined;
		if (entry.rcUsername) {
			try {
				const user = await read.getUserReader().getByUsername(entry.rcUsername);
				rcUserId = user?.id;
			} catch (error) {
				// fall through to id
			}
		}
		if (!rcUserId && entry.rcUserId) {
			try {
				const user = await read.getUserReader().getById(entry.rcUserId);
				rcUserId = user?.id;
			} catch (error) {
				// unresolvable
			}
		}
		if (!rcUserId) {
			skipped++;
			continue;
		}
		await UserMapping.persist(persistence, rcUserId, teamsUserId);
		restored++;
	}

	// Re-sync so the setting reflects what actually landed.
	await syncMappingBackupAsync({ read, app });
	return { restored, skipped };
};
