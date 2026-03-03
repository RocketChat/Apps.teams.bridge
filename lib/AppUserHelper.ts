import {
    IHttp,
    IPersistence,
    IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import { AppSetting } from "../config/Settings";
import {
    getApplicationAccessTokenAsync,
    listTeamsUserProfilesAsync,
} from "./MicrosoftGraphApi";
import { TeamsUserProfile } from "./PersistHelper";

/**
 * Fetches all Teams user profiles from the directory and persists them for use
 * in the "Add Teams User" contextual bar.  Under the single-bot architecture
 * this replaces the old `syncAllTeamsBotUsersAsync` which also created per-user
 * dummy RC bot accounts — those are no longer needed.
 */
export const syncTeamsUserProfilesAsync = async (
    http: IHttp,
    read: IRead,
    persis: IPersistence,
): Promise<void> => {
    const aadTenantId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadTenantId)
    ).value;
    const aadClientId = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadClientId)
    ).value;
    const aadClientSecret = (
        await read
            .getEnvironmentReader()
            .getSettings()
            .getById(AppSetting.AadClientSecret)
    ).value;

    const response = await getApplicationAccessTokenAsync(
        http,
        aadTenantId,
        aadClientId,
        aadClientSecret
    );
    const appAccessToken = response.accessToken;

    const teamsUserProfiles = await listTeamsUserProfilesAsync(
        http,
        appAccessToken
    );

    for (const profile of teamsUserProfiles) {
        await TeamsUserProfile.persist(
            persis,
            profile.displayName,
            profile.givenName,
            profile.surname,
            profile.mail,
            profile.id
        );
    }
};
