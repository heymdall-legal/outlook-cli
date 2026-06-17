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

function syncPage({ syncKey, receivedAt, more }: { syncKey: string; receivedAt: string; more: boolean }) {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<Sync xmlns="AirSync:" xmlns:email="Email:" xmlns:airsyncbase="AirSyncBase:"><Collections><Collection><SyncKey>' +
    syncKey +
    "</SyncKey><Status>1</Status>" +
    (more ? "<MoreAvailable/>" : "") +
    "<Commands><Add><ServerId>2:" +
    syncKey +
    "</ServerId><ApplicationData><email:Subject>S</email:Subject><email:From>a@b.com</email:From><email:DateReceived>" +
    receivedAt +
    "</email:DateReceived><airsyncbase:Body><airsyncbase:Type>1</airsyncbase:Type><airsyncbase:Data>B</airsyncbase:Data></airsyncbase:Body></ApplicationData></Add></Commands>" +
    "</Collection></Collections></Sync>"
  );
}

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

describe("syncFolderMessages cutoff paging", () => {
  it("stops paging once a whole page is older than the cutoff", async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(wbxmlResponse(syncPage({ syncKey: "1", receivedAt: "2026-06-17T11:00:00.000Z", more: true })))
      .mockResolvedValueOnce(wbxmlResponse(syncPage({ syncKey: "2", receivedAt: "2026-06-10T11:00:00.000Z", more: true })));

    const client = new ActiveSyncClient(makeConfig(), { fetchImpl, logger: { error() {} } });
    const since = new Date("2026-06-17T00:00:00.000Z");

    const result = await client.syncFolderMessages({
      mailbox: "inbox",
      collectionId: "2",
      maxPages: 50,
      windowSize: 50,
      since,
    });

    // Page 1 (in-window) + page 2 (older, triggers stop) fetched; no third page.
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect(result).toHaveLength(2);
  });
});
