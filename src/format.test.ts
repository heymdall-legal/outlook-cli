import { describe, expect, it } from "vitest";

import { formatEmail, formatEmailList } from "./format.js";
import type { NormalizedMailboxMessage } from "./types.js";

function entry(overrides: Partial<NormalizedMailboxMessage["message"]> = {}): NormalizedMailboxMessage {
  return {
    mailbox: "inbox",
    message: {
      id: "2:1",
      subject: "Subject",
      from: "a@b.com",
      to: ["c@d.com"],
      cc: [],
      receivedAt: "2026-06-17T09:00:00.000Z",
      preview: "preview",
      bodyText: "Full body",
      bodyHtml: "",
      ...overrides,
    },
  };
}

describe("formatEmailList", () => {
  it("renders a title-only row with the · separator", () => {
    const out = formatEmailList([entry()]);
    expect(out).toBe("inbox · 2026-06-17T09:00:00.000Z · 2:1 · Subject · a@b.com\n");
  });

  it("includes the body when requested", () => {
    const out = formatEmailList([entry()], { includeBody: true });
    expect(out).toContain("Full body");
  });

  it("reports an empty result", () => {
    expect(formatEmailList([])).toBe("No emails found in the requested window.\n");
  });
});

describe("formatEmail", () => {
  it("renders headers and falls back to preview when bodyText is empty", () => {
    const out = formatEmail(entry({ bodyText: "" }));
    expect(out).toContain("Subject: Subject");
    expect(out).toContain("preview");
  });
});
