import { afterEach, describe, expect, it, vi } from "vitest";
import type { PluginLogger } from "../types.js";
import { getOutlookEmails } from "./client.js";
import { classifyHttpError, createDiscoveryCandidates } from "./activesync/discovery.js";
import { outlookPlugin } from "./index.js";
import { mapMailboxMessageToMemoryDraft } from "./mapping.js";
import { filterMessagesSince, filterRecentMessages } from "./messages.js";
import type { NormalizedMailboxMessage } from "./types.js";

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

function createLogger(): PluginLogger {
  const logger: PluginLogger = {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    child: vi.fn(() => logger),
  };

  return logger;
}

describe("filterRecentMessages", () => {
  it("keeps only messages inside the first-run lookback window", () => {
    const items: NormalizedMailboxMessage[] = [
      {
        mailbox: "inbox",
        message: {
          id: "inside",
          subject: "Inside window",
          from: "boss@example.com",
          to: ["me@example.com"],
          cc: [],
          receivedAt: "2026-05-24T10:00:00.000Z",
          preview: "",
          bodyText: "Inside",
          bodyHtml: "",
        },
      },
      {
        mailbox: "inbox",
        message: {
          id: "outside",
          subject: "Outside window",
          from: "boss@example.com",
          to: ["me@example.com"],
          cc: [],
          receivedAt: "2026-05-20T10:00:00.000Z",
          preview: "",
          bodyText: "Outside",
          bodyHtml: "",
        },
      },
    ];

    const filtered = filterRecentMessages(items, 1, new Date("2026-05-25T00:00:00.000Z"));
    expect(filtered.map((item) => item.message.id)).toEqual(["inside"]);
  });
});

describe("filterMessagesSince", () => {
  it("keeps only messages at or after lastRunAt", () => {
    const items: NormalizedMailboxMessage[] = [
      {
        mailbox: "inbox",
        message: {
          id: "older",
          subject: "Older",
          from: "boss@example.com",
          to: ["me@example.com"],
          cc: [],
          receivedAt: "2026-05-24T07:59:59.000Z",
          preview: "",
          bodyText: "Older",
          bodyHtml: "",
        },
      },
      {
        mailbox: "outbox",
        message: {
          id: "newer",
          subject: "Newer",
          from: "me@example.com",
          to: ["boss@example.com"],
          cc: [],
          receivedAt: "2026-05-24T08:00:00.000Z",
          preview: "",
          bodyText: "Newer",
          bodyHtml: "",
        },
      },
    ];

    const filtered = filterMessagesSince(items, "2026-05-24T08:00:00.000Z");
    expect(filtered.map((item) => item.message.id)).toEqual(["newer"]);
  });
});

describe("mapMailboxMessageToMemoryDraft", () => {
  it("maps inbox and outbox messages into deterministic MemoryDraft values", () => {
    const draft = mapMailboxMessageToMemoryDraft({
      mailbox: "outbox",
      message: {
        id: "message-2",
        subject: "Re: Launch checklist",
        from: "me@example.com",
        to: ["partner@example.com"],
        cc: ["manager@example.com"],
        receivedAt: "2026-04-29T19:00:00.000Z",
        preview: "Sharing the final checklist.",
        bodyText: "Here is the final launch checklist.",
        bodyHtml: "<p>Here is the final launch checklist.</p>",
      },
    });

    expect(draft).toEqual({
      external_id: "message-2",
      type: "outlook.email",
      title: "Re: Launch checklist",
      text: [
        "From: me@example.com",
        "To: partner@example.com",
        "Cc: manager@example.com",
        "Received: 2026-04-29T19:00:00.000Z",
        "",
        "Body:",
        "Here is the final launch checklist.",
      ].join("\n"),
      datetime: "2026-04-29T19:00:00.000Z",
    });
  });

  it("uses a deterministic fallback external_id when message id is absent", () => {
    const draft = mapMailboxMessageToMemoryDraft({
      mailbox: "inbox",
      message: {
        subject: "Ship Q2 plan",
        from: "ceo@example.com",
        to: ["me@example.com"],
        cc: [],
        receivedAt: "2026-04-29T20:15:00.000Z",
        preview: "Need the final version tonight.",
        bodyText: "",
        bodyHtml: "",
      },
    });

    expect(draft).not.toBeNull();
    expect(draft!.external_id).toBe("inbox:2026-04-29T20:15:00.000Z:ceo@example.com:Ship Q2 plan");
    expect(draft!.type).toBe("outlook.email");
    expect(draft!.text).toContain("Body:\nNeed the final version tonight.");
  });

  it("returns null for messages with invalid timestamps", () => {
    const draft = mapMailboxMessageToMemoryDraft({
      mailbox: "inbox",
      message: {
        id: "bad-date",
        subject: "Broken",
        from: "sender@example.com",
        to: ["me@example.com"],
        cc: [],
        receivedAt: "",
        preview: "",
        bodyText: "ignore me",
        bodyHtml: "",
      },
    });

    expect(draft).toBeNull();
  });
});

describe("createDiscoveryCandidates", () => {
  it("builds default, autodiscover, and mail host candidates", () => {
    expect(
      createDiscoveryCandidates({
        endpoint: "",
        host: "exchange.example.com",
        email: "me@example.com",
      }),
    ).toEqual([
      "https://exchange.example.com/Microsoft-Server-ActiveSync",
      "https://autodiscover.exchange.example.com/Microsoft-Server-ActiveSync",
      "https://mail.exchange.example.com/Microsoft-Server-ActiveSync",
    ]);
  });
});

describe("classifyHttpError", () => {
  it("classifies common ActiveSync discovery failures", () => {
    expect(classifyHttpError({ status: 401 })).toBe("bad-credentials");
    expect(classifyHttpError({ status: 403 })).toBe("device-blocked");
    expect(classifyHttpError({ status: 404 })).toBe("endpoint-not-found");
  });
});

describe("getOutlookEmails", () => {
  it("configures the ActiveSync client and returns newest-first filtered messages", async () => {
    const now = new Date();
    const older = new Date(now);
    older.setUTCDate(older.getUTCDate() - 4);
    const newer = new Date(now);
    newer.setUTCDate(newer.getUTCDate() - 1);
    const mailboxMessages: NormalizedMailboxMessage[] = [
      {
        mailbox: "inbox",
        message: {
          id: "older",
          subject: "Older",
          from: "boss@example.com",
          to: ["me@example.com"],
          cc: [],
          receivedAt: older.toISOString(),
          preview: "",
          bodyText: "Older",
          bodyHtml: "",
        },
      },
      {
        mailbox: "outbox",
        message: {
          id: "newer",
          subject: "Newer",
          from: "me@example.com",
          to: ["boss@example.com"],
          cc: [],
          receivedAt: newer.toISOString(),
          preview: "",
          bodyText: "Newer",
          bodyHtml: "",
        },
      },
    ];
    const getMailboxMessages = vi.fn(async (): Promise<NormalizedMailboxMessage[]> => mailboxMessages);

    const items = await getOutlookEmails({
      email: "me@example.com",
      username: "user",
      password: "pass",
      host: "exchange.example.com",
      days: 2,
    }, {
      createClient: () => ({
        getMailboxMessages,
      }),
    });

    expect(getMailboxMessages).toHaveBeenCalledWith({ days: 2 });
    expect(items.map((item) => item.message.id)).toEqual(["newer"]);
  });
});

describe("outlookPlugin", () => {
  it("uses initialLookbackDays on first run", async () => {
    const getOutlookEmailsSpy = vi.spyOn(await import("./client.js"), "getOutlookEmails").mockResolvedValueOnce([
      {
        mailbox: "inbox",
        message: {
          id: "message-1",
          subject: "Ship Q2 plan",
          from: "ceo@example.com",
          to: ["me@example.com"],
          cc: [],
          receivedAt: "2026-05-24T20:15:00.000Z",
          preview: "",
          bodyText: "Please send the plan tonight.",
          bodyHtml: "",
        },
      },
    ]);

    const drafts = await outlookPlugin.run({
      config: {
        email: "me@example.com",
        username: "user",
        password: "pass",
        host: "exchange.example.com",
        initialLookbackDays: 1,
      },
      logger: createLogger(),
      signal: new AbortController().signal,
      lastRunAt: null,
    });

    expect(getOutlookEmailsSpy).toHaveBeenCalledWith({
      email: "me@example.com",
      username: "user",
      password: "pass",
      host: "exchange.example.com",
      days: 1,
    });
    expect(drafts).toHaveLength(1);
    expect(drafts[0]).toMatchObject({
      external_id: "message-1",
      type: "outlook.email",
    });
  });

  it("filters fetched messages by lastRunAt on incremental runs", async () => {
    vi.spyOn(await import("./client.js"), "getOutlookEmails").mockResolvedValueOnce([
      {
        mailbox: "inbox",
        message: {
          id: "older",
          subject: "Older",
          from: "boss@example.com",
          to: ["me@example.com"],
          cc: [],
          receivedAt: "2026-05-24T07:59:59.000Z",
          preview: "",
          bodyText: "Older",
          bodyHtml: "",
        },
      },
      {
        mailbox: "outbox",
        message: {
          id: "newer",
          subject: "Newer",
          from: "me@example.com",
          to: ["boss@example.com"],
          cc: [],
          receivedAt: "2026-05-24T08:00:00.000Z",
          preview: "",
          bodyText: "Newer",
          bodyHtml: "",
        },
      },
    ]);

    const drafts = await outlookPlugin.run({
      config: {
        email: "me@example.com",
        username: "user",
        password: "pass",
        host: "exchange.example.com",
        initialLookbackDays: 1,
      },
      logger: createLogger(),
      signal: new AbortController().signal,
      lastRunAt: "2026-05-24T08:00:00.000Z",
    });

    expect(drafts.map((draft) => draft.external_id)).toEqual(["newer"]);
  });
});
