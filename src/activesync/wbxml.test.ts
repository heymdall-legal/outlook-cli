import { describe, expect, it } from "vitest";

import { decodeWbxml, encodeWbxml } from "./wbxml.js";

describe("ItemOperations WBXML", () => {
  it("round-trips an ItemOperations Fetch request with nested AirSync/AirSyncBase elements", () => {
    const xml =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<ItemOperations xmlns="ItemOperations:" xmlns:airsync="AirSync:" xmlns:airsyncbase="AirSyncBase:">' +
      "<Fetch><Store>Mailbox</Store>" +
      "<airsync:CollectionId>2</airsync:CollectionId>" +
      "<airsync:ServerId>2:42</airsync:ServerId>" +
      "<Options><airsyncbase:BodyPreference><airsyncbase:Type>1</airsyncbase:Type>" +
      "<airsyncbase:TruncationSize>200000</airsyncbase:TruncationSize></airsyncbase:BodyPreference></Options>" +
      "</Fetch></ItemOperations>";

    const decoded = decodeWbxml(encodeWbxml(xml));

    expect(decoded).toContain("<ItemOperations");
    expect(decoded).toContain("<Store>Mailbox</Store>");
    expect(decoded).toContain("<CollectionId>2</CollectionId>");
    expect(decoded).toContain("<ServerId>2:42</ServerId>");
    expect(decoded).toContain("<Type>1</Type>");
  });
});
