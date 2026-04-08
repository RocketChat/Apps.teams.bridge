import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiUploadToDriveUrl } from '../Const';
import type { UploadFileResponse } from './types';

export const uploadFileToOneDriveAsync = async (
	http: IHttp,
	fileName: string,
	fileMIMEType: string,
	fileSize: number,
	content: Buffer,
	userAccessToken: string,
): Promise<UploadFileResponse | undefined> => {
	if (fileSize < 4096000) {
		const url = getGraphApiUploadToDriveUrl(fileName);

		const httpRequest: IHttpRequest = {
			headers: {
				Authorization: `Bearer ${userAccessToken}`,
				'Content-Type': fileMIMEType,
			},
			content: content as any as string,
		};

		const response = await http.put(url, httpRequest);

		if ([HttpStatusCode.CREATED, HttpStatusCode.OK].includes(response.statusCode)) {
			const responseBody = response.data;
			if (responseBody === undefined) {
				throw new Error('Upload file to one drive failed!');
			}

			const result: UploadFileResponse = {
				driveItemId: responseBody.id,
				fileName: responseBody.name,
				size: responseBody.size,
			};

			return result;
		}
		throw new Error(`Upload file to one drive failed with http status code ${response.statusCode}.`);
	} else {
		// TODO: implement resumable upload
		return undefined;
	}
};
