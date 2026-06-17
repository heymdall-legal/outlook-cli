export interface NormalizedMessage {
  id?: string;
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  receivedAt: string;
  preview: string;
  bodyText: string;
  bodyHtml: string;
}

export type MailboxKind = "inbox" | "sent";

export interface NormalizedMailboxMessage {
  mailbox: MailboxKind;
  message: NormalizedMessage;
}

export interface ExchangeConfig {
  email: string;
  username: string;
  password: string;
  host: string;
}

export interface BaseExchangeConfig {
  endpoint: string;
  verbose: boolean;
  deviceId: string;
  deviceType: string;
  userAgent: string;
  protocolVersion: string;
  deviceModel: string;
  deviceImei: string;
  deviceFriendlyName: string;
  deviceOs: string;
  deviceOsLanguage: string;
  devicePhoneNumber: string;
  deviceMobileOperator: string;
}

export type MailConfig = Partial<BaseExchangeConfig> &
  ExchangeConfig & {
    days: number;
  };

export interface ActiveSyncClientConfig extends BaseExchangeConfig, ExchangeConfig {
  days?: number;
  policyKey?: string;
}
