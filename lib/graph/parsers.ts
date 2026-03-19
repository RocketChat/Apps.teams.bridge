import { MessageType, MessageContentType, ThreadType } from './types';

export const parseMessageType = (messageType: string, eventDetail?: any): MessageType | undefined => {
    if (!messageType) {
        return undefined;
    }
    if (messageType === 'message') {
        return MessageType.Message;
    } else {
        if (eventDetail) {
            if (eventDetail['@odata.type'] === '#microsoft.graph.membersAddedEventMessageDetail') {
                return MessageType.SystemAddMembers;
            }
            if (eventDetail['@odata.type'] === '#microsoft.graph.membersDeletedEventMessageDetail') {
                return MessageType.SystemRemoveMembers;
            }
        }
    }
    return undefined;
};

export const parseMessageContentType = (messageContentType: string): MessageContentType | undefined => {
    if (!messageContentType) {
        return undefined;
    }
    if (messageContentType === 'html') {
        return MessageContentType.Html;
    }
    return undefined;
};

export const parseThreadType = (threadType: string): ThreadType | undefined => {
    if (!threadType) {
        return undefined;
    }
    if (threadType === 'group') {
        return ThreadType.Group;
    }
    if (threadType === 'oneOnOne') {
        return ThreadType.OneOnOne;
    }
    return undefined;
};
