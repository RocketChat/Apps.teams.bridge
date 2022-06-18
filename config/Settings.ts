import { ISetting, SettingType} from '@rocket.chat/apps-engine/definition/settings';

export enum AppSetting {
    AadTenantId = 'teamsbridge_aad_tenant_id',
    AadClientId = 'teamsbridge_aad_client_id',
    AadClientSecret = 'teamsbridge_aad_client_secret',
}

export const settings: Array<ISetting> = [
    {
        id: AppSetting.AadTenantId,
        public: false,
        type: SettingType.STRING,
        packageValue: '',
        i18nLabel: AppSetting.AadTenantId,
        required: true,
    },
    {
        id: AppSetting.AadClientId,
        public: false,
        type: SettingType.STRING,
        packageValue: '',
        i18nLabel: AppSetting.AadClientId,
        required: true,
    },
    {
        id: AppSetting.AadClientSecret,
        public: false,
        type: SettingType.STRING,
        packageValue: '',
        i18nLabel: AppSetting.AadClientSecret,
        required: true,
    },
];
