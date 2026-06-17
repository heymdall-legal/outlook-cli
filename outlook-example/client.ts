import { ActiveSyncClient } from "./activesync/client.js";
import { filterRecentMessages } from "./messages.js";
import type { MailConfig, NormalizedMailboxMessage } from "./types.js";

interface OutlookClientLike {
  getMailboxMessages(opts: { days?: number }): Promise<NormalizedMailboxMessage[]>;
}

interface GetOutlookEmailsDependencies {
  createClient?: (config: ConstructorParameters<typeof ActiveSyncClient>[0]) => OutlookClientLike;
}

export async function getOutlookEmails(
  mailConfig: MailConfig,
  dependencies: GetOutlookEmailsDependencies = {},
) {
  const config = {
    verbose: false,
    endpoint: "",
    deviceId: "NodeiPhoneClient001",
    deviceType: "iPhone",
    userAgent: "Apple-iPhone14C3/1704.10",
    protocolVersion: "14.1",
    deviceModel: "iPhone",
    deviceImei: "000000000000000",
    deviceFriendlyName: "Node iPhone Mail Reader",
    deviceOs: "iOS 18.0",
    deviceOsLanguage: "en-us",
    devicePhoneNumber: "0000000000",
    deviceMobileOperator: "Unknown",
    ...mailConfig,
  };

  const createClient =
    dependencies.createClient ??
    ((clientConfig: ConstructorParameters<typeof ActiveSyncClient>[0]) =>
      new ActiveSyncClient(clientConfig, {
        logger: console,
      }));

  const client = createClient(config);
  const allMessages = await client.getMailboxMessages({
    days: config.days,
  });
  const recentMessages = filterRecentMessages(allMessages, config.days);

  return recentMessages.sort((left, right) => {
    const leftValue = new Date(left.message.receivedAt).valueOf();
    const rightValue = new Date(right.message.receivedAt).valueOf();
    return rightValue - leftValue;
  });
}
