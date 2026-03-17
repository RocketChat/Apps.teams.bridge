import {
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from "@rocket.chat/apps-engine/definition/metadata";

const KEY = "RecentActivity";
const DEFAULT_WINDOW_MS = 3000;
const STALE_CUTOFF_MS = 30000;

export type ActivityKind = "create" | "edit" | "interaction";

export interface RecentActivityModel {
    timestamp: number;
    rcUserId: string;
    teamsThreadId: string;
    kind: ActivityKind;
}

export const RecentActivity = {
    async set(options: {
        read: IRead;
        persistence: IPersistence;
        rcUserId: string;
        teamsThreadId: string;
        kind: ActivityKind;
    }): Promise<string | null> {
        const { persistence, rcUserId, teamsThreadId, kind,read } = options;
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                `${rcUserId}:${teamsThreadId}:${kind}`,
            ),
        ];

        const data: RecentActivityModel = {
            timestamp: Date.now(),
            rcUserId,
            teamsThreadId,
            kind,
        };

        try {
            return persistence.updateByAssociations(associations, data, true);
        } catch (error) {
            console.error("[MS Teams RecentActivity] Failed to set activity:", error);
            return null;
        }
    },

    async isRecent(options: {
        read: IRead;
        rcUserId: string;
        teamsThreadId: string;
        kind: ActivityKind;
        windowMs?: number;
    }): Promise<boolean> {
        const {
            read,
            rcUserId,
            teamsThreadId,
            kind,
            windowMs = DEFAULT_WINDOW_MS,
        } = options;
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                `${rcUserId}:${teamsThreadId}:${kind}`,
            ),
        ];

        try {
            const results = (await read
                .getPersistenceReader()
                .readByAssociations(associations)) as RecentActivityModel[];

            if (!results || results.length === 0) {
                return false;
            }

            const record = results[0];
            return Date.now() - record.timestamp < windowMs;
        } catch (error) {
            console.error("[MS Teams RecentActivity] Failed to read activity:", error);
            return false;
        }
    },

    async delete(options: {
        persistence: IPersistence;
        rcUserId: string;
        teamsThreadId: string;
        kind: ActivityKind;
    }): Promise<void> {
        const { persistence, rcUserId, teamsThreadId, kind } = options;
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                `${rcUserId}:${teamsThreadId}:${kind}`,
            ),
        ];

        try {
            await persistence.removeByAssociations(associations);
        } catch (error) {
            console.error("[MS Teams RecentActivity] Failed to delete activity:", error);
        }
    },

    async deleteStale(read: IRead, persistence: IPersistence): Promise<number> {
        let count = 0;
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                KEY,
            ),
        ];

        try {
            const results = (await read
                .getPersistenceReader()
                .readByAssociations(associations)) as RecentActivityModel[];

            if (!results || results.length === 0) {
                return 0;
            }

            const cutoff = Date.now() - STALE_CUTOFF_MS;

            for (const r of results) {
                if (
                    r.timestamp < cutoff &&
                    r.rcUserId &&
                    r.teamsThreadId &&
                    r.kind
                ) {
                    await persistence.removeByAssociations([
                        new RocketChatAssociationRecord(
                            RocketChatAssociationModel.MISC,
                            KEY,
                        ),
                        new RocketChatAssociationRecord(
                            RocketChatAssociationModel.MISC,
                            `${r.rcUserId}:${r.teamsThreadId}:${r.kind}`,
                        ),
                    ]);
                    count++;
                }
            }
        } catch (error) {
            console.error(
                "[MS Teams RecentActivity] Failed during stale cleanup:",
                error,
            );
        }

        return count;
    },
};
