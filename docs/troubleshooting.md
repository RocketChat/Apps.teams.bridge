# Troubleshooting

This guide covers common issues encountered when setting up or using the Microsoft Teams Bridge, along with their root causes and solutions.

---

## Setup Issues

### Setup verification fails

**Symptom:** `/teamsbridge-setup-verification` returns *"FAILED"*.

**Root cause:** The Azure credentials (Tenant ID, Client ID, or Client Secret) are incorrect or the API permissions are not properly configured.

**Solution:**
1. Navigate to **Administration → Apps → Microsoft Teams Bridge → Settings**
2. Verify each value matches what is shown in the Azure Portal:
   - **Tenant ID** → Azure app **Overview → Directory (tenant) ID**
   - **Client ID** → Azure app **Overview → Application (client) ID**
   - **Client Secret** → Azure app **Certificates & Secrets → Value** (not the Secret ID)
3. In Azure, verify that all required API permissions have **admin consent granted** (green checkmark)
4. Run `/teamsbridge-setup-verification` again

---

### "App Bot User is not logged into Teams"

**Symptom:** Setup verification shows the Azure connection works but the bot is not logged in.

**Root cause:** The admin has not completed `/teamsbridge-login-app-user`.

**Solution:**
1. Run `/teamsbridge-login-app-user` in any Rocket.Chat room
2. Click the **"Login Teams"** button
3. Complete the OAuth login in the browser
4. Run `/teamsbridge-setup-verification` to confirm

---

### OAuth callback fails or shows an error

**Symptom:** After clicking "Login Teams", the browser shows an error instead of a success message.

**Possible causes & solutions:**

| Cause | Solution |
|-------|----------|
| Redirect URI mismatch | Ensure the redirect URI in Azure matches the app's `GET auth` endpoint URL exactly |
| Rocket.Chat not accessible | If running locally, use a tunnel service (e.g. Ngrok) and update the Proxy URL setting |
| Permissions not consented | Go to Azure **API permissions** and click **Grant admin consent** |
| Client secret expired | Create a new secret in Azure and update the app setting |

---

## Messaging Issues

### Messages not relaying from RC to Teams

**Symptom:** Messages sent in a bridged Rocket.Chat room do not appear in Microsoft Teams.

**Diagnostic steps:**

1. **Check bridge status:** Run `/teamsbridge-status`
   - If ❌ → Add the App Bot User to the room
2. **Check bot login:** Run `/teamsbridge-setup-verification`
   - If ⚠️ → Run `/teamsbridge-login-app-user`
3. **Check room type:** Only private channels, private teams, and private discussions are supported
4. **Check app logs:** Navigate to **Administration → Apps → Microsoft Teams Bridge → Logs** for errors

---

### Messages not arriving from Teams to RC

**Symptom:** Messages sent in the linked Teams chat do not appear in Rocket.Chat.

**Root cause:** Webhook subscriptions may have expired or failed.

**Solution:**
1. Verify the App Bot User is logged in (`/teamsbridge-setup-verification`)
2. Run `/teamsbridge-resubscribe-messages` to re-register webhooks
3. Verify your Rocket.Chat instance is accessible from the internet
   - Microsoft Graph needs to reach the webhook endpoint
   - If using a tunnel service, ensure it is running
   - If behind a firewall, configure the **Proxy URL** setting

---

### "success: true" in logs but nothing happens

**Symptom:** App logs show `{"success": true}` but no visible effect (user not added, message not sent).

**Root cause:** The `success: true` response is from the Rocket.Chat Apps Engine confirming the **handler method ran without crashing** — it does **not** mean the Microsoft Teams API call succeeded. The actual operation likely failed silently because:
- The App Bot User has no token (`/teamsbridge-login-app-user` not run)
- The room is not bridged (bot not in room)

**Solution:** Complete the missing setup steps and try again.

---

## Room & Member Issues

### "This room is not bridged to Microsoft Teams"

**Symptom:** Slash commands or action buttons show this error.

**Root cause:** The App Bot User is not a member of the room.

**Solution:** Add `microsoftteamsbridge.bot` (or whatever the bot username is) to the room.

---

### Teams user not added after clicking "Add users"

**Symptom:** The modal closes after clicking "Add users" but the Teams user does not appear in the room.

**Possible causes:**
1. The room is not bridged → add the bot to the room first
2. The App Bot User is not logged in → run `/teamsbridge-login-app-user`
3. The Teams user already exists in the linked chat

**Diagnostic:** Check **Administration → Apps → Microsoft Teams Bridge → Logs** for detailed error output.

---

### "Adding a Teams Bot user only supported for private channels…"

**Symptom:** Error when running `/teamsbridge-add-user` or clicking "Add Teams user".

**Root cause:** You're in a public channel or direct message.

**Solution:** Switch to a private channel, private team, or private discussion.

---

## Token & Subscription Issues

### Client secret expired

**Symptom:** The bridge stops working entirely — setup verification fails.

**Solution:**
1. Go to Azure Portal → your app → **Certificates & Secrets**
2. Create a new client secret
3. Copy the value immediately
4. Update the **Microsoft Entra Client Secret** setting in the app
5. Save changes and run `/teamsbridge-setup-verification`

---

### Webhook subscriptions expired

**Symptom:** Outbound messages (RC → Teams) work, but inbound messages (Teams → RC) stop arriving.

**Root cause:** Microsoft Graph webhook subscriptions expire after a maximum of 1 hour. The app auto-renews them every 30 minutes, but this can fail if the bot's token becomes invalid.

**Solution:**
1. Verify the bot is logged in: `/teamsbridge-setup-verification`
2. Re-register subscriptions: `/teamsbridge-resubscribe-messages`

---

### Infinite retry loop / HTTP 412 errors

**Symptom:** The Rocket.Chat server becomes unresponsive and logs show repeated HTTP 412 errors.

**Root cause:** The app is attempting operations (adding users, managing subscriptions) without a valid bot token, causing the request to fail and retry indefinitely.

**Solution:**
1. Restart the Rocket.Chat server if it is unresponsive
2. Complete all setup steps — especially `/teamsbridge-login-app-user`
3. Verify with `/teamsbridge-setup-verification`

---

## Quick Diagnostic Checklist

When something isn't working, run through this checklist:

| # | Check | Command |
|---|-------|---------|
| 1 | Is the room bridged? | `/teamsbridge-status` |
| 2 | Is setup correct? | `/teamsbridge-setup-verification` |
| 3 | Is the bot logged in? | `/teamsbridge-setup-verification` (check output) |
| 4 | Are webhooks active? | `/teamsbridge-resubscribe-messages` |
| 5 | Is RC accessible from internet? | Check Proxy URL setting or tunnel service |
| 6 | Is the client secret valid? | Check expiration in Azure Portal |
| 7 | Check app logs | **Administration → Apps → Logs** |
