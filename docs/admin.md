**Organization Admin Guide**
----

# Setup the TeamsBridge App

To enable Rocket.Chat and Microsoft Teams collaboration for your organization with TeamsBridge App, there are some actions required for the organization admin. Please make sure you have access for both your organization's `Rocket.Chat admin account` and `Microsoft Teams admin account`.

## Install the TeamsBridge App

- On Rocket.Chat Marketplace, find the Teams Bridge App and install.
- Navigate to App Admin Page, scroll to APIs, and copy `GET auth endpoint URL`. This endpoint URL will be used in following steps. If it’s a localhost url, tunnel service such as Ngrok is required to expose the localhost port to the internet.

## Setup an Azure Active Directory App for your organization

- Login to [Microsoft Azure Portal](portal.azure.com) with Microsoft Teams admin account of your organization. Find and select `Azure Active Directory` service with the search box.
- Select `Add` => `App registration`.
- Give your AAD App a name. Select `Accounts in this organizational directory only` for Supported account type. Select `Web` as Redirect URI type and paste `GET auth endpoint URL` (copied in previous step) of the TeamsBridge App here as value. Click Register to complete the registration.
- An AAD App will be created after a while. Copy `Application (client) ID` and `Directory (tenant) ID` showed on the Overview page. Those IDs will be used in following steps. Then, select `Certificates & Secret` blade.
- Select `New client secret`. Give the secret a meaningful name and click add. A client secret will be created. Copy the `Client Secret` value which will be used in following steps.
- Navigate to the `API permissions` blade. Select `Add a permission`, add the set of required permissions, and click `grant admin consent for org`.
   - **TODO: YUQING - figure the minimum required permissions set and add a required permissions list here**

## Configure the TeamsBridge App

- Navigate to the App Admin Page again, scroll to Settings. Paste `Directory (tenant) ID`, `Application (client) ID`, and `Client Secret` copied in previous step to corresponding box and click `Save changes`.

## Verify the TeamsBridge App setup correctly

- You're almost there! Now run slash command `teamsbridge-setup-verification` to verify whether you setup the TeamsBridge App correctly for your organization.

### Trouble Shooting Guide

If the slash command `teamsbridge-setup-verification` prompts nagative result, you can try the following steps.

- Make sure in Rocket.Chat App Info Settings page you see the correct value for `Directory (tenant) ID`, `Application (client) ID`, and `Client Secret`.
- Make sure you gives all required permission on `Microsoft Azure Portal - Azure Active Directory - API permissions` blade.

# Setup Microsoft Teams access for Rocket.Chat users in your organization

Document under development.
