# Organization Admin Guide

## Setup the TeamsBridge App

To enable Rocket.Chat and Microsoft Teams collaboration for your organization with TeamsBridge App, there are some actions required for the organization admin. Please make sure you have access for both your organization's `Rocket.Chat admin account` and `Microsoft Teams admin account`.

### 1. Install the TeamsBridge App

- On Rocket.Chat Marketplace, find the Teams Bridge App and install.
- Navigate to App Admin Page, scroll to APIs, and copy `GET auth endpoint URL`. This endpoint URL will be used in following steps. If it’s a localhost url, tunnel service such as Ngrok is required to expose the localhost port to the internet.

### 2. Setup an Microsoft Entra ID App for your organization

- Login to [Microsoft Azure Portal](portal.azure.com) with Microsoft Teams admin account of your organization. Find and select `Microsoft Entra ID` service with the search box.
- Select `Add` => `App registration`.
- Give your AAD App a name. Select `Accounts in this organizational directory only` for Supported account type. Select `Web` as Redirect URI type and paste `GET auth endpoint URL` (copied in previous step) of the TeamsBridge App here as value. Click Register to complete the registration.
- An AAD App will be created after a while. Copy `Application (client) ID` and `Directory (tenant) ID` showed on the Overview page. Those IDs will be used in following steps. Then, select `Certificates & Secret` blade.
- Select `New client secret`. Give the secret a meaningful name and click add. A client secret will be created. Copy the `Client Secret` value which will be used in following steps.
- Navigate to the `API permissions` blade. Select `Add a permission`, add the set of required permissions, and click `grant admin consent for org`.
   - **TODO: YUQING - figure the minimum required permissions set and add a required permissions list here**

### 3. Configure the TeamsBridge App

- Navigate to the App Admin Page again, scroll to Settings. Paste `Directory (tenant) ID`, `Application (client) ID`, and `Client Secret` copied in previous step to corresponding box and click `Save changes`.

### 4. Verify the TeamsBridge App setup correctly

- You're almost there! Now run slash command `teamsbridge-setup-verification` to verify whether you setup the TeamsBridge App correctly for your organization.

#### Trouble Shooting Guide

If the slash command `teamsbridge-setup-verification` prompts nagative result, you can try the following steps.

- Make sure in Rocket.Chat App Info Settings page you see the correct value for `Directory (tenant) ID`, `Application (client) ID`, and `Client Secret`.
- Make sure you gives all required permission on `Microsoft Azure Portal - Microsoft Entra ID - API permissions` blade.

## Setup Microsoft Teams access for Rocket.Chat users in your organization

Microsoft Teams data access is controlled by Microsoft Authorization policies. TeamsBridge App is a tool to centralize a user's access instead of extending their access. Rocket.Chat users will have to authorize the TeamsBridge App with their Microsoft Teams access in order to enable themselves collaborate with their colleagues on Microsoft Teams. People who use Microsoft Teams will be able to chat with Rocket.Chat users as if they were also on Microsoft Teams. Rocket.Chat users will be able to chat with Microsoft Teams people while staying in Rocket.Chat.

Organization admin will have to grant Microsoft Teams access to Rocket.Chat users by providing Microsoft Teams accounts so that they can authorize the TeamsBridge App to start the collaboration. Both `Guest Access account` and `general Teams account` work for the TeamsBridge App. We highly recommand setting up `Guest Access account` for Rocket.Chat users who do not have a Teams account in the organization to minimize the cost.

### Setup Guest Access account

- `Guest Access account` grants `Guest Access` to someone who doesn't have a school or work account with the organization. For details, see [Guest access in Microsoft Teams](https://docs.microsoft.com/en-us/microsoftteams/guest-access).

- Organization Admin need to make sure `Guest Access` is turned `ON` for the organization. For details, see [Teams guest access settings](https://docs.microsoft.com/en-us/microsoft-365/solutions/collaborate-as-team?view=o365-worldwide#teams-guest-access-settings).

- Once the `Guest Access` is turned `ON`, organization Admin can invite guest to the organization. This can be done either via [Teams Admin Center](https://admin.teams.microsoft.com/) or [Azure Portal](https://portal.azure.com). For details, see [Teams Admin Center solution](https://support.microsoft.com/en-us/office/add-guests-to-a-team-in-teams-fccb4fa6-f864-4508-bdde-256e7384a14f?ui=en-us&rs=en-us&ad=us) and [Azure Portal solution](https://docs.microsoft.com/en-us/azure/active-directory/external-identities/b2b-quickstart-add-guest-users-portal).

- `Guest Access` is available for limited Microsoft 365 subscription types. But the cost for a `Guest Access` is significantly lower than `general Teams account`, For details, see [Licensing for guest access](https://docs.microsoft.com/en-us/microsoftteams/guest-access#licensing-for-guest-access).

### Setup general Teams account

- To create a `general Teams account`, organization admin need to add a user via [Microsoft 365 Admin Center](https://admin.microsoft.com/Adminportal/Home#/homepage) with their Microsoft admin account. For details, see [Microsoft official document](https://docs.microsoft.com/en-us/microsoft-365/admin/add-users/add-new-employee?view=o365-worldwide).

- Then, organization admin need to assign a Teams license to the newly created user. For details, see [Microsoft official document](https://docs.microsoft.com/en-us/microsoftteams/user-access#using-the-microsoft-365-admin-center).

- The cost for a `general Teams account` depends on the plan your organization choose. For details, see [Microsoft official document](https://www.microsoft.com/en-us/microsoft-teams/compare-microsoft-teams-options?activetab=pivot%3aprimaryr1).
