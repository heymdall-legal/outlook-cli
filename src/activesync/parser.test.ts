import { describe, expect, it } from "vitest";

import { decodeWbxml, encodeWbxml } from "./wbxml.js";
import {
  buildFolderSyncRequestXml,
  findInboxFolder,
  findOutboxFolder,
  normalizeSyncMessages,
  parseFolderSyncXml,
  parseSyncXml,
} from "./parser.js";

describe("FolderSync parsing", () => {
  const sample =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<FolderSync xmlns="FolderHierarchy:"><Status>1</Status><SyncKey>1</SyncKey>' +
    "<Changes>" +
    "<Add><ServerId>2</ServerId><ParentId>0</ParentId><DisplayName>Inbox</DisplayName><Type>2</Type></Add>" +
    "<Add><ServerId>5</ServerId><ParentId>0</ParentId><DisplayName>Sent Items</DisplayName><Type>5</Type></Add>" +
    "</Changes></FolderSync>";

  it("extracts folders and finds inbox + sent", () => {
    const parsed = parseFolderSyncXml(sample);
    expect(parsed.status).toBe("1");
    expect(parsed.folders).toHaveLength(2);
    expect(findInboxFolder(parsed.folders)?.serverId).toBe("2");
    expect(findOutboxFolder(parsed.folders)?.serverId).toBe("5");
  });
});

describe("Sync message normalization", () => {
  it("maps a plain-text email into NormalizedMessage", () => {
    const sample =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<Sync xmlns="AirSync:"><Collections><Collection><SyncKey>2</SyncKey><Status>1</Status>' +
      "<Commands><Add><ServerId>2:1</ServerId><ApplicationData>" +
      "<Subject>Hi</Subject><From>a@b.com</From><DateReceived>2026-06-17T09:00:00.000Z</DateReceived>" +
      "<Type>1</Type><Data>Body</Data>" +
      "</ApplicationData></Add></Commands></Collection></Collections></Sync>";
    const parsed = parseSyncXml(sample);
    const [message] = normalizeSyncMessages(parsed.messages);
    expect(message.id).toBe("2:1");
    expect(message.subject).toBe("Hi");
    expect(message.bodyText).toBe("Body");
  });
});

describe("WBXML round-trip", () => {
  it("encodes and decodes a FolderSync request without losing structure", () => {
    const xml = buildFolderSyncRequestXml("0");
    const decoded = decodeWbxml(encodeWbxml(xml));
    expect(decoded).toContain("<FolderSync");
    expect(decoded).toContain("<SyncKey>0</SyncKey>");
  });
});
