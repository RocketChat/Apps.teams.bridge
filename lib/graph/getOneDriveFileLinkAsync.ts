import type { IHttp, IHttpRequest } from '@rocket.chat/apps-engine/definition/accessors';
import { HttpStatusCode } from '@rocket.chat/apps-engine/definition/accessors';

import { getGraphApiOneDriveFileLinkUrl } from '../Const';

export const getOneDriveFileLinkAsync = async (http: IHttp, oneDriveItemId: string, userAccessToken: string): Promise<string> => {
	const url = getGraphApiOneDriveFileLinkUrl(oneDriveItemId);

	const httpRequest: IHttpRequest = {
		headers: {
			'Content-Type': 'application/json',
			Authorization: `Bearer ${userAccessToken}`,
		},
	};

	const response = await http.get(url, httpRequest);

	if (response.statusCode === HttpStatusCode.OK) {
		const responseBody = response.data;
		if (responseBody === undefined) {
			throw new Error('Get one drive file link failed!');
		}

		const result: string = responseBody.webUrl as string;
		return result;
	}
	throw new Error(`Get one drive file link failed with http status code ${response.statusCode}.`);
};
