import { describe, expect, it } from "vitest";

import {
  buildItemOperationsFetchXml,
  normalizeSyncMessages,
  parseItemOperationsXml,
} from "./parser.js";

describe("buildItemOperationsFetchXml", () => {
  it("includes the collection id, server id, and a plain-text body preference", () => {
    const xml = buildItemOperationsFetchXml({ collectionId: "2", serverId: "2:42" });
    expect(xml).toContain("<airsync:CollectionId>2</airsync:CollectionId>");
    expect(xml).toContain("<airsync:ServerId>2:42</airsync:ServerId>");
    expect(xml).toContain("<airsyncbase:Type>1</airsyncbase:Type>");
  });
});

describe("parseItemOperationsXml", () => {
  it("parses a successful Fetch response into a normalizable message", () => {
    const xml =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<ItemOperations xmlns="ItemOperations:"><Status>1</Status><Response><Fetch>' +
      "<Status>1</Status><CollectionId>2</CollectionId><ServerId>2:42</ServerId>" +
      "<Properties><Subject>Hello</Subject><From>a@b.com</From>" +
      "<DateReceived>2026-06-17T09:00:00.000Z</DateReceived>" +
      "<Body><Type>1</Type><Data>Body text</Data></Body></Properties>" +
      "</Fetch></Response></ItemOperations>";

    const result = parseItemOperationsXml(xml);
    expect(result.status).toBe("1");
    expect(result.message?.serverId).toBe("2:42");

    const [normalized] = normalizeSyncMessages([result.message!]);
    expect(normalized.subject).toBe("Hello");
    expect(normalized.bodyText).toBe("Body text");
  });

  it("returns a null message when the item is absent", () => {
    const xml =
      '<?xml version="1.0" encoding="utf-8"?>' +
      '<ItemOperations xmlns="ItemOperations:"><Status>1</Status><Response></Response></ItemOperations>';
    const result = parseItemOperationsXml(xml);
    expect(result.message).toBeNull();
  });
});
