import { describe, expect, it, vi } from "vitest";

import { OutlookClient, buildClientConfig } from "./client.js";
import type { ResolvedConfig } from "./config.js";
import type { NormalizedMailboxMessage } from "./types.js";

const config: ResolvedConfig = {
  email: "u@e.com",
  username: "u@e.com",
  host: "e.com",
  password: "secret",
};

function message(mailbox: "inbox" | "sent", receivedAt: string): NormalizedMailboxMessage {
  return {
    mailbox,
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

describe("buildClientConfig", () => {
  it("merges device defaults with user config", () => {
    const built = buildClientConfig(config);
    expect(built.email).toBe("u@e.com");
    expect(built.deviceType).toBe("iPhone");
    expect(built.protocolVersion).toBe("14.1");
  });
});

describe("OutlookClient.listEmails", () => {
  it("filters by mailbox and sorts newest-first", async () => {
    const now = Date.now();
    const recent = new Date(now - 60 * 60 * 1000).toISOString();
    const older = new Date(now - 2 * 60 * 60 * 1000).toISOString();
    const stub = {
      resolveMailFolders: vi.fn(),
      getMailboxMessages: vi.fn().mockResolvedValue([
        message("inbox", older),
        message("sent", recent),
        message("inbox", recent),
      ]),
      getMessageById: vi.fn(),
    };

    const client = new OutlookClient(config, { createClient: () => stub });
    const result = await client.listEmails({ hours: 24, mailbox: "inbox" });

    expect(result.map((e) => e.mailbox)).toEqual(["inbox", "inbox"]);
    expect(result[0].message.receivedAt).toBe(recent);
  });
});

describe("OutlookClient.getEmail", () => {
  it("delegates to getMessageById", async () => {
    const found = message("inbox", new Date().toISOString());
    const stub = {
      resolveMailFolders: vi.fn(),
      getMailboxMessages: vi.fn(),
      getMessageById: vi.fn().mockResolvedValue(found),
    };
    const client = new OutlookClient(config, { createClient: () => stub });
    await expect(client.getEmail({ id: "2:1" })).resolves.toBe(found);
    expect(stub.getMessageById).toHaveBeenCalledWith({ serverId: "2:1" });
  });
});
