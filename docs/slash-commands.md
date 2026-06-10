# Slash Commands Reference

The Microsoft Teams Bridge provides slash commands for setup, authentication, room management, and diagnostics. All are prefixed with `/teamsbridge-`.

---

## Command Summary

| Command | Who Can Run | Purpose |
|---------|:-----------:|---------|
| `/teamsbridge-setup-verification` | Admin | Verify Azure + bot token configuration |
| `/teamsbridge-login-app-user` | Admin | Log in the App Bot User to Teams |
| `/teamsbridge-logout-app-user` | Admin | Log out the App Bot User from Teams |
| `/teamsbridge-login-teams` | Any user | Link your personal Teams account |
| `/teamsbridge-logout-teams` | Any user | Unlink your personal Teams account |
| `/teamsbridge-add-user` | Any user | Add Teams users to a bridged room |
| `/teamsbridge-view-members` | Any user | View Teams members in linked chat |
| `/teamsbridge-status` | Any user | Check if room is bridged |
| `/teamsbridge-resubscribe-messages` | Admin | Re-register webhook subscriptions |

---

## Admin Commands

### `/teamsbridge-setup-verification`

**Permission:** `manage-apps`

Verifies the Azure AD connection and checks whether the App Bot User has a valid access token.

**What it checks:**
1. Obtains an application token using Tenant ID, Client ID, and Client Secret
2. Checks if the App Bot User has a valid delegated token
3. Verifies the token by calling Microsoft Graph `GET /me`

**Outputs:**

| Output | Meaning |
|--------|---------|
| *"...verification PASSED!"* | ✅ Azure credentials and bot login are valid |
| *"Azure AD connection is verified, but the App Bot User is not logged in…"* | ⚠️ Run `/teamsbridge-login-app-user` |
| *"...verification FAILED!"* | ❌ Check Tenant ID, Client ID, or Client Secret |

---

### `/teamsbridge-login-app-user`

**Permission:** `manage-apps` + `admin` role

Generates an OAuth2 login URL for the App Bot User. Required during initial setup — without this, the bridge **cannot relay messages**.

- If already logged in → *"No need to login again."*
- If not → displays a **"Login Teams"** button

---

### `/teamsbridge-logout-app-user`

**Permission:** `manage-apps` + `admin` role

Logs out the App Bot User — revokes tokens, deletes subscriptions, clears mappings.

> **Warning:** The bridge stops relaying messages until you run `/teamsbridge-login-app-user` again.

---

### `/teamsbridge-resubscribe-messages`

**Permission:** `manage-apps` + `admin` role

Re-registers webhook subscriptions with Microsoft Graph. Run this if inbound messages (Teams → RC) stop arriving.

- If bot not logged in → prompts you to log in first
- If logged in → re-subscribes and confirms success

---

## User Commands

### `/teamsbridge-login-teams`

**Permission:** Any user

Links your personal Microsoft Teams account via OAuth2. **Optional** — provides identity-preserved messaging.

- If already logged in → *"No need to login again."*
- If not → displays a **"Login Teams"** button

**After logging in:** Messages appear under your own Teams identity instead of the bot's.

---

### `/teamsbridge-logout-teams`

**Permission:** Any user

Unlinks your personal Teams account. Your messages will still be relayed via the bot's bridged message format.

---

### `/teamsbridge-add-user`

**Permission:** Any user

Opens a contextual bar with live search from the Azure AD directory to add Teams users to the linked Teams chat.

**Requirements:**
- Room must be **bridged** (bot is a member)
- Room must be a **private channel, private team, or private discussion**

**Errors:**

| Error | Cause |
|-------|-------|
| *"This room is not bridged…"* | Bot not in room |
| *"Adding a Teams Bot user only supported for private channels…"* | Wrong room type |

---

### `/teamsbridge-view-members`

**Permission:** Any user

Opens a contextual bar showing all members of the linked Microsoft Teams chat. Same room type requirements as `/teamsbridge-add-user`.

---

### `/teamsbridge-status`

**Permission:** Any user

Checks if the current room is bridged.

| Output | Meaning |
|--------|---------|
| ✅ *"This room is **bridged**…"* | Active |
| ❌ *"This room is **not bridged**…"* | Bot not in room |
