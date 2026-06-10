# Setting Up the Microsoft Teams Bridge

This guide walks a Rocket.Chat administrator through the complete setup process. All steps must be completed **in order**.

---

## Prerequisites

- **Rocket.Chat admin account** with `manage-apps` permission
- **Microsoft Teams admin account** (or an account with permission to register apps in [Microsoft Entra ID](https://entra.microsoft.com))
- A publicly accessible Rocket.Chat instance URL (or a tunnel service such as [Ngrok](https://ngrok.com) if running locally)

---

## Step 1 — Install the App

1. Navigate to **Administration → Marketplace** in Rocket.Chat.
2. Search for **Microsoft Teams Bridge** and install the app.
3. After installation, go to the app's **Info** page and scroll down to the **APIs** section.
4. Copy the **`GET auth` endpoint URL** — you will need it in the next step.

> **Note:** If your Rocket.Chat instance is running on `localhost`, you must use a tunnel service like Ngrok to expose the port to the internet. Microsoft Entra ID needs to reach this URL during the OAuth callback.

---

## Step 2 — Register an App in Microsoft Entra ID

1. Sign in to the [Microsoft Azure Portal](https://portal.azure.com) with your Microsoft Teams admin account.
2. Search for and open **Microsoft Entra ID**.
3. Navigate to **App registrations → New registration**.
4. Fill in the registration form:
   - **Name:** Choose a descriptive name (e.g. `Rocket.Chat Teams Bridge`)
   - **Supported account types:** Select **"Accounts in this organizational directory only"**
   - **Redirect URI:**
     - **Platform:** Web
     - **URI:** Paste the `GET auth` endpoint URL copied in Step 1
5. Click **Register**.
6. On the app's **Overview** page, copy:
   - **Application (client) ID**
   - **Directory (tenant) ID**
7. Navigate to **Certificates & Secrets → New client secret**:
   - Give it a meaningful description
   - Choose an expiration period
   - Click **Add** and **immediately copy the secret value** (it will not be shown again)

---

## Step 3 — Configure API Permissions

Navigate to **API permissions** in the Azure app registration and add the following Microsoft Graph permissions:

### Delegated Permissions

These are requested during the OAuth consent flow. You must add **all** of them — some are used by the normal user login and some by the App Bot User login.

| Permission | Used By | Required For |
|------------|---------|-------------|
| `offline_access` | Both | Refresh token support |
| `openid` | Normal user | OpenID Connect authentication |
| `user.read` | Both | Reading the logged-in user's profile |
| `user.read.all` | Bot user | Searching users in the Azure AD directory |
| `chat.create` | Bot user | Creating new Teams chat threads |
| `chat.readwrite` | Both | Reading and writing to chats |
| `chat.readbasic` | Both | Reading basic chat metadata |
| `chatmember.read` | Bot user | Reading chat member lists |
| `chatmember.readwrite` | Bot user | Adding/removing members from Teams chats |
| `chatmessage.read` | Both | Reading chat messages |
| `chatmessage.send` | Both | Sending chat messages |
| `files.readwrite` | Both | Uploading/downloading files via OneDrive |

### Application Permissions

| Permission | Required For |
|------------|-------------|
| `user.read.all` | Searching users in the Azure AD directory |
| `chat.create` | Creating Teams chat threads |
| `chat.readbasic.all` | Reading basic metadata for all chats |
| `chat.readwrite.all` | Reading and writing all chats |
| `chatmember.read.all` | Reading members of all chats |
| `chatmember.readwrite.all` | Adding/removing members in all chats |
| `chatmessage.read.all` | Reading messages in all chats (webhook notifications) |
| `files.read.all` | Downloading shared files from Teams |

> **Important:** After adding all permissions, click **"Grant admin consent for \<your organization\>"** to approve them. All permissions must show a green **"Granted"** status. The OAuth consent screen will **reject** any scope that is not registered in the app registration.

---

## Step 4 — Configure App Settings

1. In Rocket.Chat, navigate to **Administration → Apps → Microsoft Teams Bridge → Settings**.
2. Fill in the following fields using values from Step 2:

| Setting | Value |
|---------|-------|
| **Microsoft Entra Directory (tenant) ID** | The `Directory (tenant) ID` from Azure |
| **Microsoft Entra Application (client) ID** | The `Application (client) ID` from Azure |
| **Microsoft Entra Client Secret** | The client secret value from Azure |
| **Proxy URL** *(optional)* | A publicly accessible URL to use as a proxy if Rocket.Chat is behind a firewall |

3. Click **Save changes**.

For full details on each setting, see the [Settings Reference](./settings.md).

---

## Step 5 — Log In the App Bot User

This is the most critical setup step. Without it, the bridge **cannot write** to Microsoft Teams.

1. Open any Rocket.Chat room.
2. Run the slash command:
   ```
   /teamsbridge-login-app-user
   ```
3. You will receive a message with a **"Login Teams"** button. Click it.
4. A Microsoft login page will open in your browser. Sign in with the Teams account you want the bot to use as its relay identity.
5. Approve the requested permissions when prompted.
6. If successful, the browser will display: **"Login to Teams succeed! You can close this window now."**

> **Who should run this?** Only an admin with `manage-apps` permission.
>
> **What does this do?** It generates a delegated OAuth2 access token and stores it under the App Bot User's identity. This token is used for all relay operations — sending messages on behalf of non-logged-in users, adding members to Teams threads, and managing webhook subscriptions.

---

## Step 6 — Verify the Setup

Run the verification command to confirm everything is configured correctly:

```
/teamsbridge-setup-verification
```

### Expected Results

| Result | Meaning |
|--------|---------|
| **"TeamsBridge app setup verification PASSED!"** | ✅ Everything is working — Azure connection and bot login are both verified. |
| **"Azure AD connection is verified, but the App Bot User is not logged into Teams…"** | ⚠️ Azure credentials are correct, but Step 5 was not completed. Run `/teamsbridge-login-app-user`. |
| **"TeamsBridge app setup verification FAILED!"** | ❌ Azure credentials are incorrect. Double-check Tenant ID, Client ID, and Client Secret. |

---

## What's Next?

Once setup is complete, you can:

1. **[Create a bridged room](./bridged-rooms.md)** — Add the bot to a room to start relaying messages.
2. **[Set up individual user login](./bridged-rooms.md#optional-individual-user-login)** — Optional step for identity-preserved messaging.
3. **[Explore slash commands](./slash-commands.md)** — See all available commands.
