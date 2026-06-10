# Frequently Asked Questions

## General

### What is the Microsoft Teams Bridge?

The Microsoft Teams Bridge is a Rocket.Chat app that enables real-time cross-platform communication between Rocket.Chat and Microsoft Teams. Messages, files, and member updates are relayed bidirectionally so users on each platform can collaborate without switching tools.

### Do all Rocket.Chat users need to log in to Microsoft Teams?

**No.** Individual user login is completely **optional**. The bridge works through a single App Bot User that relays messages for everyone. If a user has not logged in, their messages still appear in Teams — they are sent via the bot using a rich blockquote format that includes the user's display name.

Logging in via `/teamsbridge-login-teams` is only needed if a user wants their messages to appear under their **own** Teams identity.

### What Microsoft 365 license is needed?

The bridge requires at least one Microsoft 365 account with Microsoft Teams access to serve as the App Bot User. For individual users who want identity-preserved messaging, each would need their own Microsoft 365 account (either a full account or Guest Access).

### Does the bridge support public channels?

No. Only **private channels**, **private teams**, and **private discussions** can be bridged. Public channels and direct messages are not supported.

---

## Setup

### What permissions does the Azure app need?

See the [Setup Guide — Step 3](./setup.md#step-3--configure-api-permissions) for the full list. Key delegated permissions include `chat.readwrite`, `chatmessage.send`, and `files.readwrite`. The main application permission is `user.read.all`.

### Why do I need to run `/teamsbridge-login-app-user`?

This command generates a delegated OAuth2 token for the App Bot User. Without this token, the bridge can **read** from Azure AD (for user search) but cannot **write** — it cannot send messages, add members, or manage webhook subscriptions.

### Can I change the Teams account used by the bot?

Yes. Run `/teamsbridge-logout-app-user` to log out the current account, then run `/teamsbridge-login-app-user` and log in with a different Teams account.

### What happens if my client secret expires?

The bridge will stop working. You will need to:
1. Create a new client secret in the Azure Portal
2. Update the **Microsoft Entra Client Secret** setting in the app
3. Save changes

---

## Bridged Rooms

### How do I bridge a room?

Add the App Bot User (e.g. `microsoftteamsbridge.bot`) as a member of the room. See [Creating a Bridged Room](./bridged-rooms.md) for details.

### How do I un-bridge a room?

Remove the App Bot User from the room. The room will stop relaying messages immediately.

### Why can't I bridge a public channel or DM?

This is by design. The bridge only supports private channels, private teams, and private discussions. This ensures controlled access to the cross-platform communication.

### When is the Teams thread created?

A Teams thread is created automatically when the first message is sent in a bridged room or when the first Teams user is added.

---

## Messaging

### Why aren't my messages appearing in Teams?

Check the following in order:
1. Is the room **bridged**? Run `/teamsbridge-status` to verify.
2. Is the App Bot User **logged in**? Run `/teamsbridge-setup-verification`.
3. Are **API permissions** correctly configured and granted in Azure?

### Why don't inbound messages from Teams appear in RC?

1. Verify the App Bot User is logged in
2. Run `/teamsbridge-resubscribe-messages` to re-register webhook subscriptions
3. Check that the Rocket.Chat instance is accessible from the internet (Microsoft Graph must reach the webhook endpoint)

### What message types are supported?

Text messages, rich text/markdown, emoji, URLs, file attachments, message edits, and message deletions are all supported in both directions.

### What does the "[Bridged Message]" format look like?

When a non-logged-in RC user sends a message, it appears in Teams as:

> **[Bridged Message]**
> > **John Doe**
> > ---
> > Hello from Rocket.Chat!

---

## Authentication

### Is OAuth data stored securely?

Yes. Access tokens and refresh tokens are stored in the Rocket.Chat Apps Engine's persistence layer. Tokens are automatically refreshed before expiration and stale OAuth nonces are periodically cleaned up.

### How often are tokens refreshed?

The app runs a scheduled job every **30 minutes** to renew access tokens and webhook subscriptions for all registered users.

### Can I use Guest Access accounts?

Yes. Microsoft Teams Guest Access accounts work with the bridge. This can be a more cost-effective option for organizations where Rocket.Chat users don't have full Microsoft 365 licenses. See [Microsoft's Guest Access documentation](https://docs.microsoft.com/en-us/microsoftteams/guest-access) for details.
