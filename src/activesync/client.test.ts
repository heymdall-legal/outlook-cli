import { describe, expect, it, vi } from "vitest";

import { ActiveSyncClient } from "./client.js";
import { encodeWbxml } from "./wbxml.js";
import type { ActiveSyncClientConfig } from "../types.js";

function wbxmlResponse(xml: string) {
  const buffer = encodeWbxml(xml);
  return {
    ok: true,
    status: 200,
    arrayBuffer: async () => buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength),
  };
}

const FOLDER_SYNC_XML =
  '<?xml version="1.0" encoding="utf-8"?>' +
  '<FolderSync xmlns="FolderHierarchy:"><Status>1</Status><SyncKey>1</SyncKey><Changes>' +
  "<Add><ServerId>2</ServerId><ParentId>0</ParentId><DisplayName>Inbox</DisplayName><Type>2</Type></Add>" +
  "<Add><ServerId>5</ServerId><ParentId>0</ParentId><DisplayName>Sent Items</DisplayName><Type>5</Type></Add>" +
  "</Changes></FolderSync>";

const ITEM_OPS_XML =
  '<?xml version="1.0" encoding="utf-8"?>' +
  '<ItemOperations xmlns="ItemOperations:" xmlns:airsync="AirSync:" xmlns:airsyncbase="AirSyncBase:" xmlns:email="Email:">' +
  "<Status>1</Status><Response><Fetch><Status>1</Status>" +
  "<airsync:CollectionId>2</airsync:CollectionId><airsync:ServerId>2:42</airsync:ServerId>" +
  "<Properties><email:Subject>Hello</email:Subject><email:From>a@b.com</email:From>" +
  "<email:DateReceived>2026-06-17T09:00:00.000Z</email:DateReceived>" +
  "<airsyncbase:Body><airsyncbase:Type>1</airsyncbase:Type><airsyncbase:Data>Body text</airsyncbase:Data></airsyncbase:Body>" +
  "</Properties></Fetch></Response></ItemOperations>";

function makeConfig(): ActiveSyncClientConfig {
  return {
    email: "user@example.com",
    username: "user@example.com",
    password: "secret",
    host: "example.com",
    endpoint: "https://example.com/Microsoft-Server-ActiveSync",
    verbose: false,
    deviceId: "dev",
    deviceType: "iPhone",
    userAgent: "agent",
    protocolVersion: "14.1",
    deviceModel: "iPhone",
    deviceImei: "0",
    deviceFriendlyName: "fn",
    deviceOs: "iOS",
    deviceOsLanguage: "en-us",
    devicePhoneNumber: "0",
    deviceMobileOperator: "x",
  };
}

describe("ActiveSyncClient.getMessageById", () => {
  it("resolves folders then fetches the email from the inbox", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(wbxmlResponse(FOLDER_SYNC_XML)) // FolderSync
      .mockResolvedValueOnce(wbxmlResponse(ITEM_OPS_XML)); // ItemOperations Fetch

    const client = new ActiveSyncClient(makeConfig(), { fetchImpl, logger: { error() {} } });
    const result = await client.getMessageById({ serverId: "2:42" });

    expect(result?.mailbox).toBe("inbox");
    expect(result?.message.subject).toBe("Hello");
    expect(result?.message.bodyText).toBe("Body text");
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });
});
