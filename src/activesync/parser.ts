import type { NormalizedMessage } from "../types.js";
import { getAllTagBlocks, getFirstTagText, hasSelfClosingTag } from "./xml.js";

export interface FolderRecord {
  serverId: string;
  parentId?: string;
  displayName: string;
  type: string;
}

export interface FolderSyncResult {
  syncKey: string;
  status: string;
  folders: FolderRecord[];
}

interface ParsedMessageApplicationData {
  subject: string;
  from: string;
  to: string;
  cc: string;
  dateReceived: string;
  bodyType: string;
  bodyData: string;
  bodyPreview: string;
}

export interface ParsedSyncMessage {
  serverId: string;
  applicationData: ParsedMessageApplicationData;
}

export interface SyncResult {
  syncKey: string;
  status: string;
  moreAvailable: boolean;
  messages: ParsedSyncMessage[];
}

export function parseFolderSyncXml(xml: string): FolderSyncResult {
  const addBlocks = getAllTagBlocks(xml, "Add");

  return {
    syncKey: getFirstTagText(xml, "SyncKey"),
    status: getFirstTagText(xml, "Status"),
    folders: addBlocks.map((block) => ({
      serverId: getFirstTagText(block, "ServerId"),
      parentId: getFirstTagText(block, "ParentId"),
      displayName: getFirstTagText(block, "DisplayName"),
      type: getFirstTagText(block, "Type"),
    })),
  };
}

export function findInboxFolder(folders: FolderRecord[]): FolderRecord | null {
  return (
    folders.find((folder) => folder.type === "2") ??
    folders.find((folder) => folder.displayName?.trim().toLowerCase() === "inbox") ??
    folders.find((folder) => folder.serverId?.toLowerCase().includes("inbox")) ??
    null
  );
}

export function findOutboxFolder(folders: FolderRecord[]): FolderRecord | null {
  return (
    folders.find((folder) => folder.type === "5") ??
    folders.find((folder) => {
      const name = folder.displayName?.trim().toLowerCase();
      return name === "sent items" || name === "sent";
    }) ??
    folders.find((folder) => {
      const serverId = folder.serverId?.toLowerCase() ?? "";
      return serverId.includes("sent");
    }) ??
    null
  );
}

export function parseSyncXml(xml: string): SyncResult {
  const collectionBlock = getAllTagBlocks(xml, "Collection")[0] ?? "";
  const commandBlocks = [...getAllTagBlocks(collectionBlock, "Add"), ...getAllTagBlocks(collectionBlock, "Change")];

  return {
    syncKey: getFirstTagText(collectionBlock, "SyncKey"),
    status: getFirstTagText(collectionBlock, "Status") || getFirstTagText(xml, "Status"),
    moreAvailable: hasSelfClosingTag(collectionBlock, "MoreAvailable"),
    messages: commandBlocks.map((block) => parseSyncCommand(block)),
  };
}

function parseSyncCommand(xml: string): ParsedSyncMessage {
  const applicationData = getAllTagBlocks(xml, "ApplicationData")[0] ?? "";

  return {
    serverId: getFirstTagText(xml, "ServerId"),
    applicationData: {
      subject: getFirstTagText(applicationData, "Subject"),
      from: getFirstTagText(applicationData, "From"),
      to: getFirstTagText(applicationData, "To"),
      cc: getFirstTagText(applicationData, "Cc"),
      dateReceived: getFirstTagText(applicationData, "DateReceived"),
      bodyType: getFirstTagText(applicationData, "Type"),
      bodyData: getFirstTagText(applicationData, "Data"),
      bodyPreview: getFirstTagText(applicationData, "BodyPreview") || getFirstTagText(applicationData, "Preview"),
    },
  };
}

export function normalizeSyncMessages(messages: ParsedSyncMessage[]): NormalizedMessage[] {
  return messages.map(({ applicationData, serverId }) => {
    const bodyType = applicationData.bodyType;
    const bodyData = applicationData.bodyData || "";

    return {
      id: serverId,
      subject: applicationData.subject || "",
      from: applicationData.from || "",
      to: splitAddresses(applicationData.to),
      cc: splitAddresses(applicationData.cc),
      receivedAt: applicationData.dateReceived || "",
      preview: applicationData.bodyPreview || bodyData.slice(0, 160),
      bodyText: bodyType === "2" ? "" : bodyData,
      bodyHtml: bodyType === "2" ? bodyData : "",
    };
  });
}

function splitAddresses(value: string): string[] {
  if (!value) {
    return [];
  }

  return value
    .split(/[;,]/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

export function buildFolderSyncRequestXml(syncKey: string = "0"): string {
  return `<?xml version="1.0" encoding="utf-8"?><FolderSync xmlns="FolderHierarchy:"><SyncKey>${escapeXml(
    syncKey
  )}</SyncKey></FolderSync>`;
}

export function buildSyncRequestXml({
  protocolVersion = "14.1",
  syncKey,
  collectionId,
  windowSize = 50,
}: {
  protocolVersion?: string;
  syncKey: string;
  collectionId: string;
  windowSize?: number;
}): string {
  const classElement = usesLegacyClass(protocolVersion) ? "<Class>Email</Class>" : "";

  if (syncKey === "0") {
    return `<?xml version="1.0" encoding="utf-8"?><Sync xmlns="AirSync:"><Collections><Collection>${classElement}<SyncKey>${escapeXml(
      syncKey
    )}</SyncKey><CollectionId>${escapeXml(collectionId)}</CollectionId></Collection></Collections></Sync>`;
  }

  return `<?xml version="1.0" encoding="utf-8"?><Sync xmlns="AirSync:" xmlns:airsyncbase="AirSyncBase:"><Collections><Collection>${classElement}<SyncKey>${escapeXml(
    syncKey
  )}</SyncKey><CollectionId>${escapeXml(
    collectionId
  )}</CollectionId><DeletesAsMoves>0</DeletesAsMoves><GetChanges>1</GetChanges><WindowSize>${windowSize}</WindowSize><Options><MIMESupport>0</MIMESupport><MIMETruncation>0</MIMETruncation><airsyncbase:BodyPreference><airsyncbase:Type>1</airsyncbase:Type><airsyncbase:TruncationSize>200000</airsyncbase:TruncationSize></airsyncbase:BodyPreference></Options></Collection></Collections></Sync>`;
}

function usesLegacyClass(protocolVersion: string): boolean {
  return ["2.5", "12.0", "12.1"].includes(protocolVersion);
}

function escapeXml(value: string): string {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function buildItemOperationsFetchXml({
  protocolVersion = "14.1",
  collectionId,
  serverId,
  truncationSize = 200000,
}: {
  protocolVersion?: string;
  collectionId: string;
  serverId: string;
  truncationSize?: number;
}): string {
  void protocolVersion;

  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<ItemOperations xmlns="ItemOperations:" xmlns:airsync="AirSync:" xmlns:airsyncbase="AirSyncBase:">' +
    "<Fetch><Store>Mailbox</Store>" +
    `<airsync:CollectionId>${escapeXml(collectionId)}</airsync:CollectionId>` +
    `<airsync:ServerId>${escapeXml(serverId)}</airsync:ServerId>` +
    "<Options><airsyncbase:BodyPreference>" +
    "<airsyncbase:Type>1</airsyncbase:Type>" +
    `<airsyncbase:TruncationSize>${truncationSize}</airsyncbase:TruncationSize>` +
    "</airsyncbase:BodyPreference></Options>" +
    "</Fetch></ItemOperations>"
  );
}

export interface ItemOperationsResult {
  status: string;
  message: ParsedSyncMessage | null;
}

export function parseItemOperationsXml(xml: string): ItemOperationsResult {
  const topStatus = getFirstTagText(xml, "Status");
  const fetchBlock = getAllTagBlocks(xml, "Fetch")[0] ?? "";

  if (!fetchBlock) {
    return { status: topStatus, message: null };
  }

  const fetchStatus = getFirstTagText(fetchBlock, "Status") || topStatus;
  const properties = getAllTagBlocks(fetchBlock, "Properties")[0] ?? "";
  const serverId = getFirstTagText(fetchBlock, "ServerId");

  if (!serverId && !properties) {
    return { status: fetchStatus, message: null };
  }

  return {
    status: fetchStatus,
    message: {
      serverId,
      applicationData: {
        subject: getFirstTagText(properties, "Subject"),
        from: getFirstTagText(properties, "From"),
        to: getFirstTagText(properties, "To"),
        cc: getFirstTagText(properties, "Cc"),
        dateReceived: getFirstTagText(properties, "DateReceived"),
        bodyType: getFirstTagText(properties, "Type"),
        bodyData: getFirstTagText(properties, "Data"),
        bodyPreview: getFirstTagText(properties, "Preview"),
      },
    },
  };
}
