import { randomBytes } from 'crypto';

import type { IHttp, IModify, IPersistence, IRead } from '@rocket.chat/apps-engine/definition/accessors';
import { RoomType } from '@rocket.chat/apps-engine/definition/rooms';
import type { IUser } from '@rocket.chat/apps-engine/definition/users';

import type { TeamsBridgeApp } from '../../TeamsBridgeApp';
import { DefaultTeamName } from '../Const';
import { mapTeamsMessageToRocketChatMessage, sendRocketChatMessageInRoomAsync } from '../MessageHelper';
import { MessageMapping, RecentActivity, Room, UploadMapping, UserMapping } from '../PersistHelper';
import { getChatThreadWithMembersAsync, getMessageWithResourceStringAsync, MessageType, ThreadType } from '../graph';
import { getSenderUser } from './getSender';
import type { InBoundNotification } from './handleInboundNotificationAsync';
import { PreventRegistry } from '../PreventRegistry';

export const handleInboundMessageCreatedAsync = async (
	userAccessToken: string,
	inBoundNotification: InBoundNotification,
	read: IRead,
	modify: IModify,
	http: IHttp,
	persis: IPersistence,
	app: TeamsBridgeApp,
): Promise<void> => {
	const receiverRocketChatUserId = inBoundNotification.receiverRocketChatUserId;
	const resourceString = inBoundNotification.resourceString;
	const getMessageResponse = await getMessageWithResourceStringAsync(http, resourceString, userAccessToken);

	if (getMessageResponse.messageType) {
		const appUser = await read.getUserReader().getAppUser();
		const storedMessageMap = await MessageMapping.findByTeamsMessageId(read, getMessageResponse.messageId);
		if (storedMessageMap?.rocketChatMessageId) {
			// IMPORTANT!!!!!
			// An echo message. Should skip. Else this will create a loop.
			return;
		}

		// --- Adaptive delay race mitigation ---
		const fromUserTeamsId = getMessageResponse.fromTeamsUser.id;
		if (fromUserTeamsId) {
			const fromUserRocketChatUser = await UserMapping.findByTeamsUserId(read, fromUserTeamsId);
			if (fromUserRocketChatUser) {
				const isRecent = await RecentActivity.isRecent({
					read,
					rcUserId: fromUserRocketChatUser.rocketChatUserId,
					teamsThreadId: getMessageResponse.threadId,
					kind: 'create',
				});
				if (isRecent) {
					await new Promise((r) => setTimeout(r, 1000));
					const recheck = await MessageMapping.findByTeamsMessageId(read, getMessageResponse.messageId);
					await RecentActivity.delete({
						persistence: persis,
						rcUserId: fromUserRocketChatUser.rocketChatUserId,
						teamsThreadId: getMessageResponse.threadId,
						kind: 'create',
					});
					if (recheck?.rocketChatMessageId) {
						// confirmed echo
						app.getLogger().debug(`Skipping echo message for Teams message ${getMessageResponse.messageId}`);
						return;
					}
				}
			}
		}

		let roomRecord = await Room.findByTeamsThreadId(read, getMessageResponse.threadId);
		if (!roomRecord) {
			if (getMessageResponse.messageType !== MessageType.Message) {
				// Only create room for real message
				return;
			}

			// Handle thread created in Teams scenario
			// Get thread and members info
			const threadInfo = await getChatThreadWithMembersAsync(http, getMessageResponse.threadId, userAccessToken);

			// Build a room with thread info
			const userReader = read.getUserReader();
			const notificationReceiverUser = await userReader.getById(receiverRocketChatUserId);

			const topic = `${DefaultTeamName}_${randomBytes(4).toString('hex').slice(0, 8)}`;

			const creator = modify.getCreator();
			const roomBuilder = creator.startRoom();
			roomBuilder.setCreator(notificationReceiverUser);
			if (threadInfo.type) {
				if (threadInfo.type === ThreadType.OneOnOne) {
					roomBuilder.setType(RoomType.DIRECT_MESSAGE).setSlugifiedName(`dm_${notificationReceiverUser.id}`);
				} else if (threadInfo.type === ThreadType.Group) {
					roomBuilder.setType(RoomType.PRIVATE_GROUP).setDisplayName(topic).setSlugifiedName(topic);
				} else {
					throw new Error(`Unsupported thread type ${threadInfo.type} found for Teams thread ${threadInfo.threadId}`);
				}

				const teamsMemberIds = threadInfo.memberIds;
				if (!teamsMemberIds || teamsMemberIds.length === 0) {
					throw new Error(`No members found for Teams thread ${threadInfo.threadId}`);
				}

				// Add thread members to the room
				let madeFirstMemberOwner = false;
				for (const teamsMemberId of teamsMemberIds) {
					const rocketChatUser = await UserMapping.findByTeamsUserId(read, teamsMemberId);
					if (rocketChatUser) {
						const user = await userReader.getById(rocketChatUser.rocketChatUserId);
						roomBuilder.addMemberToBeAddedByUsername(user.username);
						if (!madeFirstMemberOwner && user.id !== appUser?.id) {
							roomBuilder.setCreator(user);
							madeFirstMemberOwner = true;
						}
					} else {
						// Under single-bot arch there are no dummy users. Teams-only members
						// who have no RC registration are not added to the RC room.
						console.log(`No RC user found for Teams member ${teamsMemberId}, skipping room membership.`);
					}
				}
			} else {
				throw new Error(`No thread type found for Teams thread ${threadInfo.threadId}`);
			}

			const roomId = await creator.finish(roomBuilder);
			console.log(`Room ${roomId} created for incoming message!`);

			// Persist room record
			await Room.persist(read, persis, roomId, threadInfo.threadId);

			roomRecord = await Room.findByTeamsThreadId(read, getMessageResponse.threadId);
			if (!roomRecord) {
				throw new Error(`Create room failed for Teams thread ${getMessageResponse.threadId}`);
			}
		}

		const room = await read.getRoomReader().getById(roomRecord.rocketChatRoomId);
		if (!room) {
			return;
		}

		// Only handle notification received by the app bot to avoid duplication
		if (receiverRocketChatUserId !== appUser?.id) {
			console.log('Skip notification for non-app user');
			return;
		}

		if (getMessageResponse.messageType === MessageType.Message) {
			const fromUserTeamsId = getMessageResponse.fromTeamsUser.id;
			if (!fromUserTeamsId) {
				// If there's no sender, stop processing
				console.error('No sender for message');
				return;
			}

			const fromUserRocketChatUser = await UserMapping.findByTeamsUserId(read, fromUserTeamsId);

			const senderUser = await getSenderUser({
				roomRecord,
				fromUserRocketChatUser,
				read,
				fromUserTeamsId,
			});

			if (!senderUser) {
				throw new Error('No user found to send the message');
			}

			// When the sender has no RC registration the app bot relays the message.
			// Prefix the message text with the Teams sender's display name so RC
			// users can see who originally sent it.
			const usesBotFallback = !fromUserRocketChatUser;

			const messageOptions = usesBotFallback ? { alias: getMessageResponse.fromTeamsUser.displayName ?? getMessageResponse.fromTeamsUser.id } : undefined;
			const message = await mapTeamsMessageToRocketChatMessage({
				getMessageResponse,
				accessToken: userAccessToken,
				room,
				sender: senderUser,
				http,
				modify,
				read,
				uploadFiles: true,
				app,
				persistence: persis,
				messageOptions,
			});

			const persistUploadsAsync = async () => {
				const uploadIds = await message.uploadCallback();
				return uploadIds.map((uploadIdMap) =>
					UploadMapping.persist({
						persistence: persis,
						rocketchatUploadId: uploadIdMap.rocketChat,
						teamsAttachmentId: uploadIdMap.teams,
						teamsMessageId: getMessageResponse.messageId,
						teamsThreadId: getMessageResponse.threadId,
						relayedByAppUser: usesBotFallback,
					}),
				);
			};

			const persistUploadPromises = await persistUploadsAsync();
			await Promise.all(persistUploadPromises);

			if (message.text === '') {
				// File message, no text content
				return;
			}

			const rocketChatMessageId = await sendRocketChatMessageInRoomAsync(message.text, senderUser, room, modify, read, messageOptions);

			await MessageMapping.persist({
				persistence: persis,
				rocketChatMessageId,
				teamsMessageId: getMessageResponse.messageId,
				teamsThreadId: getMessageResponse.threadId,
				relayedByAppUser: usesBotFallback,
			});
		} else if (getMessageResponse.messageType === MessageType.SystemAddMembers) {
			const memberToAddTeamsIds = getMessageResponse.memberIds;
			if (!memberToAddTeamsIds || memberToAddTeamsIds.length === 0) {
				console.error('Empty members Id list for add members.');
				return;
			}

			for (const memberToAddTeamsId of memberToAddTeamsIds) {
				// Echo prevention: skip if this add was initiated from RC side
				const captured = await PreventRegistry.capture(persis, `member-add:${getMessageResponse.threadId}:${memberToAddTeamsId}`);
				if (captured) {
					continue;
				}

				let userToAdd: IUser | undefined = undefined;

				// First, try find whether there's a real Rocket.Chat user for this Teams user to add
				const rocketChatUser = await UserMapping.findByTeamsUserId(read, memberToAddTeamsId);
				if (rocketChatUser) {
					userToAdd = await read.getUserReader().getById(rocketChatUser.rocketChatUserId);
				} else {
					// Under single-bot arch there are no dummy users. Teams members without
					// a registered RC account are not added to the RC room.
					console.log(`No RC user found for Teams member ${memberToAddTeamsId}, skipping room membership.`);
					continue;
				}

				const updater = modify.getUpdater();
				const roomBuilder = await updater.room(room.id, room.creator);

				if (!userToAdd) {
					console.error('Could not add Teams bot user to room!');
					console.error(`Dummy user with Teams ID ${memberToAddTeamsId} not found after try sync all Teams bot users!`);
					continue;
				}

				roomBuilder.addMemberToBeAddedByUsername(userToAdd.username);
				await updater.finish(roomBuilder);
			}
		} else if (getMessageResponse.messageType === MessageType.SystemRemoveMembers) {
			// APP Engine does not support removing users from rooms?
		} else {
			console.log('Unsupported message type.');
		}
	} else {
		console.log('Unsupported message type.');
	}
};
