import { describe, expect, it } from "vitest";

import { filterMessagesWithinHours } from "./messages.js";
import type { NormalizedMailboxMessage } from "./types.js";

function entry(receivedAt: string): NormalizedMailboxMessage {
  return {
    mailbox: "inbox",
    message: {
      id: receivedAt,
      subject: "s",
      from: "f",
      to: [],
      cc: [],
      receivedAt,
      preview: "",
      bodyText: "",
      bodyHtml: "",
    },
  };
}

describe("filterMessagesWithinHours", () => {
  const now = new Date("2026-06-17T12:00:00.000Z");

  it("keeps messages inside the window and drops older ones", () => {
    const messages = [
      entry("2026-06-17T11:00:00.000Z"), // 1h ago -> keep
      entry("2026-06-16T11:00:00.000Z"), // 25h ago -> drop
    ];
    const result = filterMessagesWithinHours(messages, 24, now);
    expect(result.map((e) => e.message.id)).toEqual(["2026-06-17T11:00:00.000Z"]);
  });

  it("drops messages with unparseable dates", () => {
    const result = filterMessagesWithinHours([entry("not-a-date")], 24, now);
    expect(result).toHaveLength(0);
  });
});
