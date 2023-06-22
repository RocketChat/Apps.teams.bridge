import {
  HttpStatusCode,
  IHttp,
  IModify,
  IPersistence,
  IRead,
} from "@rocket.chat/apps-engine/definition/accessors";
import {
  ApiEndpoint,
  IApiEndpointInfo,
  IApiRequest,
  IApiResponse,
} from "@rocket.chat/apps-engine/definition/api";
import { IApiResponseJSON } from "@rocket.chat/apps-engine/definition/api/IResponse";
import { IApp } from "@rocket.chat/apps-engine/definition/IApp";
import { AppSetting } from "../config/Settings";
import {
  AuthenticationEndpointPath,
  SubscriberEndpointPath,
} from "../lib/Const";
import {
  getUserAccessTokenAsync,
  getUserProfileAsync,
  listSubscriptionsAsync,
  subscribeToAllMessagesForOneUserAsync,
} from "../lib/MicrosoftGraphApi";
import {
  persistUserAccessTokenAsync,
  persistUserAsync,
  saveLoginMessageSentStatus,
} from "../lib/PersistHelper";
import { getRocketChatAppEndpointUrl } from "../lib/UrlHelper";

export class AuthenticationEndpoint extends ApiEndpoint {
  private embeddedLoginSuccessMessage = "Login to Teams succeed! You can close this window now.";
  private embeddedLoginFailureMessage = "Login to Teams failed! Please check the documentation or contact your organization admin.";

  public path = AuthenticationEndpointPath;

  constructor(app: IApp) {
    super(app);
    this.errorResponse = this.errorResponse.bind(this);
  }

  public async get(
    request: IApiRequest,
    endpoint: IApiEndpointInfo,
    read: IRead,
    modify: IModify,
    http: IHttp,
    persist: IPersistence
  ): Promise<IApiResponse> {
    if (request.query.error && !request.query.code) {
      return this.errorResponse();
    }

    try {
      const aadTenantId = await this.getAppSetting(read, AppSetting.AadTenantId);
      const aadClientId = await this.getAppSetting(read, AppSetting.AadClientId);
      const aadClientSecret = await this.getAppSetting(read, AppSetting.AadClientSecret);

      const rocketChatUserId = request.query.state;
      const accessCode = request.query.code;
      const authEndpointUrl = await getRocketChatAppEndpointUrl(this.app.getAccessors(), AuthenticationEndpointPath);

      const response = await getUserAccessTokenAsync(
        http,
        accessCode,
        authEndpointUrl,
        aadTenantId,
        aadClientId,
        aadClientSecret
      );

      const { accessToken: userAccessToken, refreshToken, expiresIn, extExpiresIn } = response;

      const teamsUserProfile = await getUserProfileAsync(http, userAccessToken);

      await Promise.all([
        persistUserAccessTokenAsync(persist, rocketChatUserId, userAccessToken, refreshToken as string, expiresIn, extExpiresIn),
        persistUserAsync(persist, rocketChatUserId, teamsUserProfile.id),
        saveLoginMessageSentStatus({ persistence: persist, rocketChatUserId, wasSent: false }),
      ]);

      const subscriptions = await listSubscriptionsAsync(http, userAccessToken);
      if (!subscriptions || subscriptions.length === 0) {
        const subscriberEndpointUrl = await getRocketChatAppEndpointUrl(this.app.getAccessors(), SubscriberEndpointPath);
        subscribeToAllMessagesForOneUserAsync(http, rocketChatUserId, teamsUserProfile.id, subscriberEndpointUrl, userAccessToken);
      }

      return this.success(this.embeddedLoginSuccessMessage);
    } catch (error) {
      return this.errorResponse();
    }
  }

  private async getAppSetting(read: IRead, settingId: AppSetting): Promise<string> {
    const setting = await read.getEnvironmentReader().getSettings().getById(settingId);
   
