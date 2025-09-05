import { IRead, IModify, IHttp, IPersistence } from '@rocket.chat/apps-engine/definition/accessors';
import { IJobContext, IProcessor } from '@rocket.chat/apps-engine/definition/scheduler';
import { handleInboundNotificationAsync } from '../lib/InboundNotificationHelper';
import { IncomingNotificationProcessorId } from '../lib/Const';
export class InboundNotificationProcessor implements IProcessor {
    id = IncomingNotificationProcessorId;
    constructor(private readonly app: any) {}
    async processor(jobContext: IJobContext, read: IRead, modify: IModify, http: IHttp, persistence: IPersistence): Promise<void> {
        await handleInboundNotificationAsync({
            app: this.app,
            read,
            modify,
            http,
            inBoundNotification: jobContext.inBoundNotification,
            persistence
        })
    }
}
