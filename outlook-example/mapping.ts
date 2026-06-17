import type { MemoryDraft } from "../types.js";
import type { NormalizedMailboxMessage } from "./types.js";

export function mapMailboxMessageToMemoryDraft(
  entry: NormalizedMailboxMessage,
): MemoryDraft | null {
  const datetime = normalizeDatetime(entry.message.receivedAt);
  if (!datetime) {
    return null;
  }

  return {
    external_id: entry.message.id ?? buildFallbackExternalId(entry, datetime),
    type: "outlook.email",
    title: entry.message.subject.trim() || "No subject",
    text: [
      `From: ${entry.message.from.trim() || "Unknown sender"}`,
      entry.message.to.length > 0 ? `To: ${entry.message.to.join(", ")}` : null,
      entry.message.cc.length > 0 ? `Cc: ${entry.message.cc.join(", ")}` : null,
      `Received: ${datetime}`,
      "",
      "Body:",
      resolveBody(entry),
    ]
      .filter((line): line is string => line !== null)
      .join("\n")
      .trim(),
    datetime,
  };
}

function buildFallbackExternalId(entry: NormalizedMailboxMessage, datetime: string): string {
  return `${entry.mailbox}:${datetime}:${entry.message.from.trim()}:${entry.message.subject.trim() || "No subject"}`;
}

function normalizeDatetime(value: string): string | null {
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return parsed.toISOString();
}

function resolveBody(entry: NormalizedMailboxMessage): string {
  const bodyText = entry.message.bodyText.trim();
  if (bodyText) {
    return bodyText;
  }

  const preview = entry.message.preview.trim();
  if (preview) {
    return preview;
  }

  return "No body";
}
