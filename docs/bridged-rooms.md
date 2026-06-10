# Creating a Bridged Room

A **bridged room** is a Rocket.Chat room that actively relays messages between Rocket.Chat and Microsoft Teams. This guide explains how to set one up and start cross-platform communication.

---

## Prerequisites

Before bridging a room, make sure the [initial setup](./setup.md) is complete:

- ✅ App is installed and configured with Azure credentials
- ✅ API permissions are granted with admin consent
- ✅ App Bot User is logged in via `/teamsbridge-login-app-user`
- ✅ Setup verification passes (`/teamsbridge-setup-verification`)

---

## Step 1 — Add the Bot to a Room

To bridge a room, invite the App Bot User (e.g. `microsoftteamsbridge.bot`) as a member:

1. Open the room you want to bridge.
2. Click the **Members** panel.
3. Click **Add Members** and search for `microsoftteamsbridge.bot` (or whatever the bot's username is).
4. Add the bot to the room.

### What happens when the bot joins

- The room is flagged as a **bridged room** (`isBridged = true`)
- A notification message is sent to the room confirming that bridging is active:
  > *"Hey, I been added to this room. So the room is now an active bridge room and I will start relaying messages between Rocket.Chat and Microsoft Teams."*
- If the App Bot User does not have a valid token (Step 5 of setup not done), room members are notified to contact an admin

### Supported Room Types

| Room Type | Can Be Bridged? |
|-----------|:-----------:|
| Private channel | ✅ |
| Private team | ✅ |
| Private discussion | ✅ |
| Public channel | ❌ |
| Direct message | ❌ |

### Deactivating a Bridge

To stop bridging a room, simply **remove the App Bot User** from the room. The room will no longer relay messages.

---

## Step 2 — Add Teams Users

Once a room is bridged, you can add Microsoft Teams users to the linked Teams thread:

### Using the Slash Command

1. Open the bridged room.
2. Run:
   ```
   /teamsbridge-add-user
   ```
3. A contextual bar (side panel) opens with a **live search** from the Azure AD directory.
4. Type a display name to search for Teams users.
5. Select one or more users from the results.
6. Click **"Add users"** to add them to the linked Teams chat.

### Using the Room Action Button

1. Open the bridged room.
2. Click the **kebab menu** (⋮) or room action menu.
3. Select **"Add Teams user"**.
4. The same contextual bar opens — search, select, and add users.

> **Note:** If you attempt to add a user in a room that is **not bridged**, you will receive an error message:
> *"This room is not bridged to Microsoft Teams. To activate bridging, add me to this room."*

---

## Step 3 — Start Messaging

Once the room is bridged and Teams users have been added, messages flow automatically:

- **RC → Teams:** Any message sent in the bridged room is relayed to the linked Teams chat.
- **Teams → RC:** Any message sent in the Teams chat is relayed back to the Rocket.Chat room.

### How Messages Appear

| Sender | Logged In? | Appears in Teams As |
|--------|:----------:|-------------------|
| RC user | Yes (via `/teamsbridge-login-teams`) | Their own Teams identity |
| RC user | No | App Bot User with a blockquote showing the sender's name |

| Sender | Has Linked RC Account? | Appears in RC As |
|--------|:----------------------:|-----------------|
| Teams user | Yes | Their RC identity |
| Teams user | No | App Bot User with the Teams user's display name as alias |

---

## Optional: Individual User Login

Individual Rocket.Chat users can **optionally** link their personal Microsoft Teams account for identity-preserved messaging:

1. Run the slash command in any room:
   ```
   /teamsbridge-login-teams
   ```
2. Click the **"Login Teams"** button in the message.
3. Complete the Microsoft OAuth login in the browser.

### Benefits of Logging In

- Messages appear under the user's **own Teams identity** instead of the bot's
- Direct identity preservation in both directions

### Without Logging In

- Messages are **still bridged** — they are relayed through the App Bot User
- In Teams, the message appears with a **[Bridged Message]** header and the sender's display name in a blockquote
- This is perfectly functional and is the default experience for most users

---

## Viewing Teams Members

To see which Microsoft Teams users are currently in the linked Teams chat:

### Using the Slash Command

```
/teamsbridge-view-members
```

### Using the Room Action Button

1. Click the kebab menu (⋮) in the bridged room.
2. Select **"View Teams members"**.

A contextual bar opens listing all members of the linked Teams thread.

---

## Checking Bridge Status

To verify whether the current room is actively bridged:

```
/teamsbridge-status
```

| Response | Meaning |
|----------|---------|
| ✅ *"This room is bridged to Microsoft Teams…"* | Room is actively bridged |
| ❌ *"This room is not bridged to Microsoft Teams…"* | Bot is not in the room — add it to activate bridging |
