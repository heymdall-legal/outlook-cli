import { ActiveSyncClient } from "./activesync/client.js";
import type { ResolvedConfig } from "./config.js";
import { filterMessagesWithinHours } from "./messages.js";
import type { ActiveSyncClientConfig, MailboxKind, NormalizedMailboxMessage } from "./types.js";

const DEVICE_DEFAULTS = {
  verbose: false,
  endpoint: "",
  deviceId: "NodeiPhoneClient001",
  deviceType: "iPhone",
  userAgent: "Apple-iPhone14C3/1704.10",
  protocolVersion: "14.1",
  deviceModel: "iPhone",
  deviceImei: "000000000000000",
  deviceFriendlyName: "Outlook CLI",
  deviceOs: "iOS 18.0",
  deviceOsLanguage: "en-us",
  devicePhoneNumber: "0000000000",
  deviceMobileOperator: "Unknown",
} as const;

export interface ActiveSyncLike {
  resolveMailFolders(): Promise<unknown>;
  getMailboxMessages(opts: { since?: Date; maxPages?: number }): Promise<NormalizedMailboxMessage[]>;
  getMessageById(opts: { serverId: string }): Promise<NormalizedMailboxMessage | null>;
}

interface OutlookClientDependencies {
  createClient?: (config: ActiveSyncClientConfig) => ActiveSyncLike;
}

export function buildClientConfig(config: ResolvedConfig): ActiveSyncClientConfig {
  return {
    ...DEVICE_DEFAULTS,
    email: config.email,
    username: config.username,
    password: config.password,
    host: config.host,
    endpoint: config.endpoint ?? DEVICE_DEFAULTS.endpoint,
    protocolVersion: config.protocolVersion ?? DEVICE_DEFAULTS.protocolVersion,
  };
}

export class OutlookClient {
  private createClient: (config: ActiveSyncClientConfig) => ActiveSyncLike;

  constructor(
    private config: ResolvedConfig,
    dependencies: OutlookClientDependencies = {},
  ) {
    this.createClient =
      dependencies.createClient ?? ((clientConfig) => new ActiveSyncClient(clientConfig, { logger: console }));
  }

  private build(): ActiveSyncLike {
    return this.createClient(buildClientConfig(this.config));
  }

  async verifyConnection(): Promise<void> {
    await this.build().resolveMailFolders();
  }

  async listEmails({
    hours,
    mailbox,
    maxPages,
  }: {
    hours: number;
    mailbox?: MailboxKind;
    maxPages?: number;
  }): Promise<NormalizedMailboxMessage[]> {
    const since = new Date(Date.now() - hours * 60 * 60 * 1000);
    const all = await this.build().getMailboxMessages({ since, maxPages });

    const withinWindow = filterMessagesWithinHours(all, hours);
    const scoped = mailbox ? withinWindow.filter((entry) => entry.mailbox === mailbox) : withinWindow;

    return scoped.sort(
      (left, right) => new Date(right.message.receivedAt).valueOf() - new Date(left.message.receivedAt).valueOf(),
    );
  }

  async getEmail({ id }: { id: string }): Promise<NormalizedMailboxMessage | null> {
    return this.build().getMessageById({ serverId: id });
  }
}
