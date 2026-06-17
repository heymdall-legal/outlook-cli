import type { ActiveSyncClientConfig, MailboxKind, NormalizedMailboxMessage } from "../types.js";
import { classifyHttpError, discoverEndpoint } from "./discovery.js";
import {
  buildFolderSyncRequestXml,
  buildSyncRequestXml,
  findInboxFolder,
  findOutboxFolder,
  normalizeSyncMessages,
  parseFolderSyncXml,
  parseSyncXml,
} from "./parser.js";
import {
  buildInitialProvisionRequestXml,
  buildProvisionAckRequestXml,
  parseProvisionResponseXml,
} from "./provision.js";
import { decodeWbxml, encodeWbxml } from "./wbxml.js";

interface LoggerLike {
  error(message: string): void;
}

interface FetchResponseLike {
  ok: boolean;
  status: number;
  arrayBuffer(): Promise<ArrayBuffer>;
}

type FetchLike = (
  input: string,
  init: {
    method: string;
    headers: Record<string, string>;
    body?: Buffer;
    redirect?: "manual";
  }
) => Promise<FetchResponseLike>;

interface MailboxMessageOptions {
  days?: number;
  maxPages?: number;
  windowSize?: number;
  reprovisionAttempts?: number;
}

interface ExecuteCommandOptions {
  policyKey?: string;
}

export class ActiveSyncClient {
  config: ActiveSyncClientConfig;
  fetchImpl: FetchLike;
  logger: LoggerLike;
  endpoint: string;
  policyKey: string;

  constructor(
    config: ActiveSyncClientConfig,
    { fetchImpl = fetch as FetchLike, logger = console }: { fetchImpl?: FetchLike; logger?: LoggerLike } = {}
  ) {
    this.config = config;
    this.fetchImpl = fetchImpl;
    this.logger = logger;
    this.endpoint = config.endpoint;
    this.policyKey = config.policyKey ?? "0";
  }

  async ensureEndpoint(): Promise<string> {
    if (this.endpoint) {
      return this.endpoint;
    }

    const discovery = await discoverEndpoint(this.config, this.fetchImpl);
    this.endpoint = discovery.endpoint;

    if (this.config.verbose) {
      this.logger.error(`Discovered endpoint ${this.endpoint} via ${JSON.stringify(discovery.attempts)}`);
    }

    return this.endpoint;
  }

  async getMailboxMessages({
    days,
    maxPages = 10,
    windowSize = 50,
    reprovisionAttempts = 0,
  }: MailboxMessageOptions = {}): Promise<NormalizedMailboxMessage[]> {
    void days;
    await this.ensureEndpoint();

    const folderSyncXml = await this.executeCommand("FolderSync", buildFolderSyncRequestXml("0"));
    const folderSync = parseFolderSyncXml(folderSyncXml);
    const inbox = findInboxFolder(folderSync.folders);
    const outbox = findOutboxFolder(folderSync.folders);

    if (folderSync.status === "142" || folderSync.status === "144") {
      if (reprovisionAttempts >= 1) {
        throw new Error(
          `FolderSync failed with ActiveSync status ${folderSync.status} after reprovision retry. Raw response: ${compactXml(
            folderSyncXml
          )}`
        );
      }

      await this.performProvisioning();
      return this.getMailboxMessages({
        days,
        maxPages,
        windowSize,
        reprovisionAttempts: reprovisionAttempts + 1,
      });
    }

    if (folderSync.status !== "1") {
      throw new Error(
        `FolderSync failed with ActiveSync status ${folderSync.status}. Raw response: ${compactXml(folderSyncXml)}`
      );
    }

    if (!inbox) {
      throw new Error(
        `Inbox folder not found in FolderSync response. Folders seen: ${JSON.stringify(
          folderSync.folders
        )}. Raw response: ${compactXml(folderSyncXml)}`
      );
    }

    const inboxMessages = await this.syncFolderMessages({
      mailbox: "inbox",
      collectionId: inbox.serverId,
      maxPages,
      windowSize,
    });
    const outboxMessages = outbox
      ? await this.syncFolderMessages({
          mailbox: "outbox",
          collectionId: outbox.serverId,
          maxPages,
          windowSize,
        })
      : [];

    return [...inboxMessages, ...outboxMessages];
  }

  async syncFolderMessages({
    mailbox,
    collectionId,
    maxPages,
    windowSize,
  }: {
    mailbox: MailboxKind;
    collectionId: string;
    maxPages: number;
    windowSize: number;
  }): Promise<NormalizedMailboxMessage[]> {
    const messages: NormalizedMailboxMessage[] = [];
    let syncKey = "0";

    for (let page = 0; page < maxPages; page += 1) {
      const requestSyncKey = syncKey;
      const xml = await this.executeCommand(
        "Sync",
        buildSyncRequestXml({
          protocolVersion: this.config.protocolVersion,
          syncKey,
          collectionId,
          windowSize,
        })
      );
      const parsed = parseSyncXml(xml);
      if (parsed.status && parsed.status !== "1") {
        throw new Error(`Sync failed with ActiveSync status ${parsed.status}. Raw response: ${compactXml(xml)}`);
      }
      syncKey = parsed.syncKey || syncKey;
      messages.push(
        ...normalizeSyncMessages(parsed.messages).map((message) => ({
          mailbox,
          message,
        }))
      );

      if (this.config.verbose) {
        this.logger.error(
          `Sync ${mailbox} page ${page + 1}: requestSyncKey=${requestSyncKey}, responseSyncKey=${parsed.syncKey}, messages=${parsed.messages.length}, moreAvailable=${parsed.moreAvailable}`
        );
      }

      if (requestSyncKey === "0" && parsed.messages.length === 0 && syncKey !== "0") {
        continue;
      }

      if (!parsed.moreAvailable) {
        break;
      }
    }

    return messages;
  }

  async performProvisioning(): Promise<void> {
    const initialXml = await this.executeCommand("Provision", buildInitialProvisionRequestXml(this.config), {
      policyKey: "0",
    });
    const initial = parseProvisionResponseXml(initialXml);

    if (initial.status !== "1" || !initial.policyKey) {
      throw new Error(`Provision phase 1 failed. Raw response: ${compactXml(initialXml)}`);
    }

    const ackXml = await this.executeCommand(
      "Provision",
      buildProvisionAckRequestXml({
        policyKey: initial.policyKey,
        status: "1",
      }),
      { policyKey: initial.policyKey }
    );
    const ack = parseProvisionResponseXml(ackXml);

    if (ack.status !== "1" || ack.policyStatus !== "1" || !ack.policyKey) {
      throw new Error(`Provision phase 2 failed. Raw response: ${compactXml(ackXml)}`);
    }

    this.policyKey = ack.policyKey;

    if (this.config.verbose) {
      this.logger.error(`Provisioning succeeded with policy key ${this.policyKey}`);
    }
  }

  async executeCommand(
    command: string,
    xml: string,
    { policyKey = this.policyKey }: ExecuteCommandOptions = {}
  ): Promise<string> {
    const endpoint = await this.ensureEndpoint();
    const body = encodeWbxml(xml);
    const url = buildCommandUrl(endpoint, command, this.config);

    let response: FetchResponseLike;

    try {
      response = await this.fetchImpl(url, {
        method: "POST",
        headers: buildCommandHeaders(this.config, policyKey),
        body,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`Network or TLS failure calling ${command}: ${message}`);
    }

    if (!response.ok) {
      const classification = classifyHttpError({ status: response.status });
      throw new Error(`ActiveSync ${command} failed with HTTP ${response.status} (${classification})`);
    }

    const responseBody = Buffer.from(await response.arrayBuffer());
    const xmlResponse = decodeWbxml(responseBody);

    if (this.config.verbose) {
      this.logger.error(`${command} raw response: ${compactXml(xmlResponse)}`);
    }

    return xmlResponse;
  }
}

function compactXml(xml: string): string {
  return xml.replace(/\s+/g, " ").trim();
}

function buildCommandUrl(endpoint: string, command: string, config: ActiveSyncClientConfig): string {
  const url = new URL(endpoint);
  url.searchParams.set("Cmd", command);
  url.searchParams.set("User", config.email);
  url.searchParams.set("DeviceId", config.deviceId);
  url.searchParams.set("DeviceType", config.deviceType);
  return url.toString();
}

function buildCommandHeaders(config: ActiveSyncClientConfig, policyKey: string = "0"): Record<string, string> {
  return {
    Authorization: `Basic ${Buffer.from(`${config.username}:${config.password}`).toString("base64")}`,
    "Content-Type": "application/vnd.ms-sync.wbxml",
    "MS-ASProtocolVersion": config.protocolVersion,
    "User-Agent": config.userAgent,
    "X-MS-PolicyKey": policyKey,
  };
}
