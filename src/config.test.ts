import { beforeEach, describe, expect, it, vi } from "vitest";

import { getConfigPath, getFileConfigPath, loadConfig, redactConfig, writeConfig } from "./config.js";

const { mockFindCredentials, mockSetPassword, mockMkdir, mockReadFile, mockWriteFile } = vi.hoisted(() => ({
  mockFindCredentials: vi.fn(),
  mockSetPassword: vi.fn(),
  mockMkdir: vi.fn(),
  mockReadFile: vi.fn(),
  mockWriteFile: vi.fn(),
}));

vi.mock("keytar", () => ({
  default: { findCredentials: mockFindCredentials, setPassword: mockSetPassword },
}));

vi.mock("node:fs/promises", () => ({
  mkdir: mockMkdir,
  readFile: mockReadFile,
  writeFile: mockWriteFile,
}));

describe("config", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("builds config paths under .config/outlook", () => {
    expect(getConfigPath("/tmp/home")).toBe("/tmp/home/.config/outlook");
    expect(getFileConfigPath("/tmp/home")).toBe("/tmp/home/.config/outlook/config.json");
  });

  it("loads config from file and keyring", async () => {
    mockReadFile.mockResolvedValueOnce(
      JSON.stringify({ email: "u@e.com", username: "u@e.com", host: "e.com" }),
    );
    mockFindCredentials.mockResolvedValueOnce([
      { account: "other", password: "ignore" },
      { account: "cli", password: "secret" },
    ]);

    await expect(loadConfig()).resolves.toEqual({
      email: "u@e.com",
      username: "u@e.com",
      host: "e.com",
      password: "secret",
      endpoint: undefined,
      protocolVersion: undefined,
    });
  });

  it("throws when required values are missing", async () => {
    mockReadFile.mockRejectedValueOnce(Object.assign(new Error("missing"), { code: "ENOENT" }));
    mockFindCredentials.mockResolvedValueOnce([]);
    await expect(loadConfig()).rejects.toThrow("Missing required configuration");
  });

  it("writes file config without the password and stores password in keyring", async () => {
    await writeConfig({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "secret" });

    expect(mockWriteFile).toHaveBeenCalledWith(
      expect.stringContaining(".config/outlook/config.json"),
      `${JSON.stringify({ email: "u@e.com", username: "u@e.com", host: "e.com" }, null, 2)}\n`,
      "utf8",
    );
    expect(mockSetPassword).toHaveBeenCalledWith("outlook", "cli", "secret");
  });

  it("redacts the password", () => {
    expect(
      redactConfig({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "secret" }),
    ).toEqual({ email: "u@e.com", username: "u@e.com", host: "e.com", password: "****" });
  });
});
