# App Settings

The Microsoft Teams Bridge has four configurable settings, accessible via **Administration → Apps → Microsoft Teams Bridge → Settings** in Rocket.Chat.

---

## Settings Reference

### Microsoft Entra Directory (tenant) ID

| Property | Value |
|----------|-------|
| **Setting ID** | `teamsbridge_aad_tenant_id` |
| **Type** | String |
| **Required** | Yes |
| **Default** | *(empty)* |

The **Directory (tenant) ID** from your Microsoft Entra ID app registration. Found on the app's **Overview** page in the Azure Portal.

This identifies your Microsoft 365 organization/tenant and is used to construct OAuth2 authorization and token endpoints.

---

### Microsoft Entra Application (client) ID

| Property | Value |
|----------|-------|
| **Setting ID** | `teamsbridge_aad_client_id` |
| **Type** | String |
| **Required** | Yes |
| **Default** | *(empty)* |

The **Application (client) ID** from your Microsoft Entra ID app registration. Found on the app's **Overview** page in the Azure Portal.

This uniquely identifies the registered application and is used in all OAuth2 flows.

---

### Microsoft Entra Client Secret

| Property | Value |
|----------|-------|
| **Setting ID** | `teamsbridge_aad_client_secret` |
| **Type** | String |
| **Required** | Yes |
| **Default** | *(empty)* |

The **Client Secret** value from your Microsoft Entra ID app registration. Found under **Certificates & Secrets** in the Azure Portal.

> **Important:** Copy the secret **value** (not the Secret ID) immediately after creation — Azure will not show it again.

> **Warning:** Client secrets have an expiration date. When a secret expires, the bridge will stop working. Create a new secret and update this setting before the current one expires.

---

### Proxy URL

| Property | Value |
|----------|-------|
| **Setting ID** | `teamsbridge_proxy_url` |
| **Type** | String |
| **Required** | No |
| **Default** | *(empty)* |

An optional publicly accessible URL to be used as a proxy if Rocket.Chat is not directly accessible from the internet. This URL replaces the Site URL during webhook registration with Microsoft Teams.

**When to use:** If your Rocket.Chat instance is behind a firewall, NAT, or reverse proxy and Microsoft Graph cannot reach it directly for webhook notifications.

**Example:** `https://proxy.mycompany.com`

---

## Where to Find These Values

1. Sign in to the [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID → App registrations**
3. Select your registered app
4. **Tenant ID** and **Client ID** are on the **Overview** page
5. **Client Secret** is under **Certificates & Secrets** (you'll need to create one if you haven't already)

For step-by-step instructions, see the [Setup Guide](./setup.md).
