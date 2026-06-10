# Overview & App Capabilities

## What Is the Microsoft Teams Bridge?

The Microsoft Teams Bridge is a Rocket.Chat app that connects Rocket.Chat rooms with Microsoft Teams chats, enabling real-time cross-platform collaboration. Users on either platform can send and receive messages without switching tools.

---

## Key Concepts

Before diving in, it helps to understand the terminology used throughout the documentation.

| Term | Definition |
|------|-----------|
| **App Bot User** | A single system-level Rocket.Chat user (e.g. `microsoftteamsbridge.bot`) created automatically when the app is installed. This is the only bot account used by the bridge — there is **no** per-Teams-user bot. It acts as the relay identity for all bridged messages. |
| **App Bot User's Linked Teams Account** | The Microsoft Teams account the App Bot User is logged into via OAuth. Also referred to as the **Bridge User**. An admin completes this login via `/teamsbridge-login-app-user`. |
| **Bridged Room** | A Rocket.Chat room where the App Bot User has been added as a member. Only bridged rooms relay messages between Rocket.Chat and Microsoft Teams. A room becomes bridged when the bot joins, and stops being bridged when the bot leaves. |
| **Logged-in RC User** | A Rocket.Chat user who has optionally linked their personal Microsoft Teams account using `/teamsbridge-login-teams`. |
| **Non-logged-in RC User** | A Rocket.Chat user who has **not** linked a personal Teams account. Their messages are still bridged — they are relayed through the App Bot User's linked Teams account using a rich blockquote format that includes their display name. |
| **Teams Thread** | The Microsoft Teams chat (group or 1:1) linked to a Rocket.Chat bridged room. Created automatically when the first message is sent or when a Teams user is added. |

---

## Supported Features

### Message Relay

| Feature | Status |
|---------|--------|
| Text messages (RC → Teams) | ✅ Supported |
| Text messages (Teams → RC) | ✅ Supported |
| Emoji rendering | ✅ Supported |
| URL link parsing & preview | ✅ Supported |
| Rich text / Markdown formatting | ✅ Supported |
| Message edits (RC → Teams) | ✅ Supported |
| Message edits (Teams → RC) | ✅ Supported |
| Message deletions (RC → Teams) | ✅ Supported |
| Message deletions (Teams → RC) | ✅ Supported |

### File Sharing

| Feature | Status |
|---------|--------|
| File uploads (RC → Teams via OneDrive) | ✅ Supported |
| File downloads (Teams → RC) | ✅ Supported |

### Member Management

| Feature | Status |
|---------|--------|
| Add Teams users to a bridged room | ✅ Supported |
| View Teams members in a linked chat | ✅ Supported |
| Live search of Azure AD directory | ✅ Supported |
| Automatic member sync on room join/leave | ✅ Supported |

### Identity & Authentication

| Feature | Status |
|---------|--------|
| OAuth2 login for the App Bot User | ✅ Supported |
| OAuth2 login for individual RC users | ✅ Supported (optional) |
| Identity-preserved messaging (logged-in users) | ✅ Supported |
| Bridged message format (non-logged-in users) | ✅ Supported |
| Automatic token refresh | ✅ Supported |
| Webhook subscription auto-renewal | ✅ Supported |

### Room Types

| Room Type | Supported |
|-----------|-----------|
| Private channels | ✅ Yes |
| Private teams | ✅ Yes |
| Private discussions | ✅ Yes |
| Public channels | ❌ No |
| Direct messages | ❌ No |

---

## How Message Relay Works

Understanding the message flow is key to understanding why individual user login is **optional**.

### Outbound (Rocket.Chat → Microsoft Teams)

```
RC user sends a message in a bridged room
  │
  ├─ Is the RC user logged in to Teams? (/teamsbridge-login-teams)
  │   │
  │   ├─ YES → Message is sent via the user's own Teams token
  │   │         (appears in Teams under their true identity)
  │   │
  │   └─ NO  → Is the App Bot User logged in? (/teamsbridge-login-app-user)
  │             │
  │             ├─ YES → Message is relayed via the App Bot User's token
  │             │         (appears in Teams using a rich blockquote format
  │             │          with the sender's display name)
  │             │
  │             └─ NO  → Message is NOT relayed to Teams
  │                       (room members are notified that admin login is required)
```

### Inbound (Microsoft Teams → Rocket.Chat)

```
Teams user sends a message in a linked Teams thread
  │
  ├─ Is the App Bot User logged in & webhook subscription active?
  │   │
  │   ├─ NO  → Messages are NOT received by Rocket.Chat
  │   │
  │   └─ YES → Message is processed by the bridge
  │         │
  │         └─ Does this Teams user have a linked RC account?
  │             │
  │             ├─ YES → Message appears in RC under their true RC identity
  │             │
  │             └─ NO  → Message appears in RC from the App Bot User
  │                       (using the Teams user's display name as a message alias)
```

> **Key takeaway:** Individual RC users do **not** need to log in for the bridge to function. If they are not logged in, their messages are relayed through the App Bot User's linked Teams account using a rich blockquote format that carries their name. Logging in is optional and provides the benefit of messages appearing under their own Teams identity.

---

## Architecture Overview

The app is built on the Rocket.Chat Apps Engine and uses the following components:

| Component | Purpose |
|-----------|---------|
| **OAuth2 Authentication** | Handles delegated and app-level token flows against Microsoft Entra ID |
| **Microsoft Graph API** | All read/write operations — messages, chats, members, subscriptions |
| **Webhook Subscriptions** | Receives real-time notifications from Teams via Microsoft Graph change notifications |
| **Scheduled Jobs** | Auto-renews tokens and subscriptions, cleans up stale data |
| **Slash Commands** | User-facing commands for login, setup, status, and management |
| **UI Action Buttons** | Context menu buttons for adding Teams users and viewing members |
| **Persistence Layer** | Stores user mappings, room mappings, message mappings, and token data |
