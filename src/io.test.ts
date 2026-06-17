import { afterEach, describe, expect, it, vi } from "vitest";

import { formatError, formatLine, parseOutputFormat, printOutput } from "./io.js";

describe("parseOutputFormat", () => {
  it("defaults to text and rejects unknown values", () => {
    expect(parseOutputFormat(undefined)).toBe("text");
    expect(parseOutputFormat("json")).toBe("json");
    expect(() => parseOutputFormat("xml")).toThrow("Unsupported output format");
  });
});

describe("formatLine", () => {
  it("returns strings unchanged and stringifies objects", () => {
    expect(formatLine("hi")).toBe("hi");
    expect(formatLine({ a: 1 })).toBe(JSON.stringify({ a: 1 }, null, 2));
  });
});

describe("formatError", () => {
  it("chains causes", () => {
    const error = new Error("top", { cause: new Error("inner") });
    expect(formatError(error)).toBe("top\nCaused by: inner\n");
  });
});

describe("printOutput", () => {
  afterEach(() => vi.restoreAllMocks());

  it("writes pretty JSON in json mode", () => {
    const write = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    printOutput({ value: { a: 1 }, format: "json" });
    expect(write).toHaveBeenCalledWith(`${JSON.stringify({ a: 1 }, null, 2)}\n`);
  });

  it("uses the text formatter in text mode", () => {
    const write = vi.spyOn(process.stdout, "write").mockReturnValue(true);
    printOutput({ value: ["x"], format: "text", textFormatter: () => "FORMATTED" });
    expect(write).toHaveBeenCalledWith("FORMATTED");
  });
});
