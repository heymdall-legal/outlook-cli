import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { mockLoadConfig, mockListEmails, mockGetEmail, mockVerify } = vi.hoisted(() => ({
  mockLoadConfig: vi.fn(),
  mockListEmails: vi.fn(),
  mockGetEmail: vi.fn(),
  mockVerify: vi.fn(),
}));

vi.mock("./config.js", () => ({
  loadConfig: mockLoadConfig,
  getFileConfigPath: () => "/tmp/config.json",
  redactConfig: (c: unknown) => c,
  writeConfig: vi.fn(),
}));

vi.mock("./client.js", () => ({
  OutlookClient: vi.fn(),
}));

import { createCli } from "./cli.js";
import { OutlookClient } from "./client.js";

describe("createCli", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(OutlookClient).mockImplementation(() => ({
      listEmails: mockListEmails,
      getEmail: mockGetEmail,
      verifyConnection: mockVerify,
    }) as never);
    mockLoadConfig.mockResolvedValue({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "p" });
    vi.spyOn(process.stdout, "write").mockReturnValue(true);
    vi.spyOn(process.stderr, "write").mockReturnValue(true);
  });

  afterEach(() => vi.restoreAllMocks());

  it("lists emails and returns exit code 0", async () => {
    mockListEmails.mockResolvedValue([]);
    const exitCode = await createCli().run(["emails", "list", "--hours", "12"]);
    expect(exitCode).toBe(0);
    expect(mockListEmails).toHaveBeenCalledWith({ hours: 12, mailbox: undefined, maxPages: 50 });
  });

  it("returns exit code 1 when get finds no email", async () => {
    mockGetEmail.mockResolvedValue(null);
    const exitCode = await createCli().run(["emails", "get", "2:1"]);
    expect(exitCode).toBe(1);
    expect(process.stderr.write).toHaveBeenCalledWith(expect.stringContaining("was not found"));
  });

  it("rejects an invalid output format with exit code 1", async () => {
    const exitCode = await createCli().run(["emails", "list", "--output", "xml"]);
    expect(exitCode).toBe(1);
    expect(process.stderr.write).toHaveBeenCalledWith(expect.stringContaining("Unsupported output format"));
  });
});
