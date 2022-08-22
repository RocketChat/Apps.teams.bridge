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

## Collaboration Experience

One of the most important goals of Rocket.Chat TeamsBridge App is to build a smooth user experience for both users on Rocket.Chat and Microsoft Teams. To achieve this, TeamsBridge App introduces a concept called `Dummy User` in Rocket.Chat world. Each Rocket.Chat `Dummy User` represents a real user on Microsoft Teams. When a Rocket.Chat user that has already `embedded login` to his Teams account sends a message to a `Dummy User`, the message will be delivered to Microsoft Teams world with the original sender's Teams account as the sender and the `Dummy User`'s corresponding Teams account as the receiver. As a result, from the Rocket.Chat users' perspective, they are just collaborating with someone on Rocket.Chat. Meanwhile, from the Teams users' perspective, they are just messaging someone on Teams. With the `Dummy User` approach, the TeamsBridge App delivers messages between Rocket.Chat and Microsoft Teams while keeping the orginal collaboration experience for users on both platforms.

### Send one on one Direct Message to a collaborator who use Microsoft Teams

To send a one on one Direct Message from Rocket.Chat to a collaborator who use Microsoft Teams, the Rocket.Chat user just need to search the `Dummy User` that represent the Teams user in Rocket.Chat client and send a message to them. The message will be delivered to Microsoft Teams world with the original sender's Teams account as the sender and the `Dummy User`'s corresponding Teams account as the receiver.

### Receive one on one Direct Message from a collaborator who use Microsoft Teams

This feature is under development and will be available soon.

### Send Message to a group chat with participant(s) who use Microsoft Teams

Currently, this is NOT a supported scenario, which is under developer and will be added soon in the future.

### Receive Message in a group chat with participant(s) who use Microsoft Teams

This feature is under development and will be available soon.

### Supported Message Types

Currently, the following message types are supported:
- Text Message
