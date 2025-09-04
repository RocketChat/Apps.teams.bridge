import { IPersistence } from "@rocket.chat/apps-engine/definition/accessors";
import {
    RocketChatAssociationModel,
    RocketChatAssociationRecord,
} from "@rocket.chat/apps-engine/definition/metadata";
import { MiscKeys } from "./PersistHelper";

/**
 * This class will help preventing a task from being executed multiple times.
 * It provides methods to set, capture, clear, and release tasks.
 * It uses a persistence layer to store the key related to each task.
 */
export class PreventRegistry {
    public static async set(
        persistence: IPersistence,
        key: string,
        value?: any,
    ): Promise<void> {
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                MiscKeys.PreventRegistry
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                key,
            ),
        ];

        console.log(`PreventRegistry set: ${key}`);

        await persistence.updateByAssociations(associations, { ...(value && { value }) }, true);
    }

    public static async capture(persistence: IPersistence, key: string) {
        console.log(`PreventRegistry capture started: ${key}`);
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                MiscKeys.PreventRegistry
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                key
            ),
        ];

        const data = await persistence.removeByAssociations(associations);

        if (!data) {
            return null;
        }
        if (data.length > 0) {
            console.log(`PreventRegistry captured: ${data.join(", ").toString()}`);
        } else {
            console.log(`PreventRegistry capture failed: ${key}`);
        }
        return data.shift() ?? null;
    }

    public static async captureMany(
        persistence: IPersistence,
        keys: string[]
    ): Promise<boolean> {
        const results = await Promise.all(keys.map(key => this.capture(persistence, key)));
        return results.some(Boolean);
    }

    public static async clear(persistence: IPersistence) {
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                MiscKeys.PreventRegistry
            ),
        ];

        await persistence.removeByAssociations(associations);
    }

    public static async release(persistence: IPersistence, key: string) {
        const associations = [
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                MiscKeys.PreventRegistry
            ),
            new RocketChatAssociationRecord(
                RocketChatAssociationModel.MISC,
                key
            ),
        ];

        await persistence.removeByAssociations(associations);
    }
}
