import type { ISetting } from '@rocket.chat/apps-engine/definition/settings';
import { SettingType } from '@rocket.chat/apps-engine/definition/settings';

export enum AppSetting {
	AadTenantId = 'teamsbridge_aad_tenant_id',
	AadClientId = 'teamsbridge_aad_client_id',
	AadClientSecret = 'teamsbridge_aad_client_secret',
	ProxyUrl = 'teamsbridge_proxy_url',
	MappingsBackup = 'teamsbridge_mappings_backup',
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
	{
		id: AppSetting.MappingsBackup,
		public: false,
		type: SettingType.STRING,
		multiline: true,
		packageValue: '',
		i18nLabel: AppSetting.MappingsBackup,
		i18nDescription: 'teamsbridge_mappings_backup_description',
		required: false,
	},
	{
		id: AppSetting.ProxyUrl,
		public: false,
		type: SettingType.STRING,
		packageValue: '',
		i18nLabel: AppSetting.ProxyUrl,
		i18nDescription: 'teamsbridge_proxy_url_description',
		required: false,
	},
];
