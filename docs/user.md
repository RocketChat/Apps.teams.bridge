# Rocket.Chat User Guide

## Embedded login

Microsoft Teams data access is controlled by Microsoft Authorization policies. Without your authorization, no one could access your Teams data or send/receive messages on-behalf-of you. In order to enable yourself collaborate with your colleagues on Microsoft Teams, you'll have to login to your Microsoft Teams account to authorize the TeamsBridge App to relay your message back and forth. Your colleagues who use Microsoft Teams will be able to chat with you as if you were also on Microsoft Teams. You'll be able to chat with them while staying in Rocket.Chat.

The required action for you to authorize the TeamsBridge App is called `Embedded Login`. You need to follow the following steps.

### 1. Make sure you have a Teams account or a Guess Account of your organization

- If you don't yet have a Teams account or a Guess Account of your organization, contact your organization admin to let them create one for you.

### 2. Get the Embedded Login link for your Rocket.Chat account

- Run the slash command `teamsbridge-login-teams` in Rocket.Chat.
- You will receive a message 'To start cross platform collaboration, you need to login to Microsoft with your Teams account or guest account You'll be able to keep using Rocket.Chat, but you'll also be able to chat with colleagues using Microsoft Teams. Login Teams'.
- The `Login Teams` contains a link, which is the Embedded Login link for your Rocket.Chat account.

### 3. Login with Microsoft in opened web page 

- Click `Login Teams` link.
- The link will redirect you to a Microsoft login web page.
- Login with your Teams account or Guest account.
- Click `Yes` if you see the login page ask authorization for specific permissions from you.
- If you successfully login to Teams and authorize the TeamsBridge App, the web pages will show `Login to Teams succeed! You can close this window now.`. You can just close the window.

## Send Direct Message to a collaborator who use Microsoft Teams

Document under development.

## Send Message to a group chat with participant(s) who use Microsoft Teams

Document under development.
