# Microsoft Teams Bridge

Rocket.Chat app to support connecting collaborators across Rocket.Chat and Microsoft Teams. Messages, files, and member updates are relayed in real time so that users on each platform can collaborate without leaving their preferred tool.

## Documentation

| Guide | Description |
|-------|-------------|
| [Overview & Capabilities](./docs/capabilities.md) | What the app does, supported features, and architecture overview |
| [Setting Up the MS Teams Bridge](./docs/setup.md) | End-to-end admin guide — Azure registration, permissions, settings, and verification |
| [Creating a Bridged Room](./docs/bridged-rooms.md) | How to activate bridging and add Microsoft Teams users |
| [Slash Commands Reference](./docs/slash-commands.md) | Complete reference for all slash commands |
| [App Settings](./docs/settings.md) | Detailed explanation of every configurable setting |
| [FAQs](./docs/faq.md) | Frequently asked questions |
| [Troubleshooting](./docs/troubleshooting.md) | Common issues, root causes, and resolution steps |

## Development

### Commands

- `rc-apps package` — Generate a packaged app file (zip) which can be installed if it compiles with TypeScript
- `rc-apps deploy` — Package and deploy; will prompt for your server URL, username, and password

### Resources

- [Rocket.Chat Apps Engine Documentation](https://rocketchat.github.io/Rocket.Chat.Apps-engine/)
- [Rocket.Chat Apps Engine Repository](https://github.com/RocketChat/Rocket.Chat.Apps-engine)
- [Example Rocket.Chat Apps](https://github.com/graywolf336/RocketChatApps)
- Community Forums
  - [App Requests](https://forums.rocket.chat/c/rocket-chat-apps/requests)
  - [App Guides](https://forums.rocket.chat/c/rocket-chat-apps/guides)
  - [Top View of Both Categories](https://forums.rocket.chat/c/rocket-chat-apps)
- [#rocketchat-apps on Open.Rocket.Chat](https://open.rocket.chat/channel/rocketchat-apps)
