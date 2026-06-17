import type { NormalizedMailboxMessage } from "./types.js";

const SEP = " · ";

export function formatEmailList(
  entries: NormalizedMailboxMessage[],
  options: { includeBody?: boolean } = {},
): string {
  if (entries.length === 0) {
    return "No emails found in the requested window.\n";
  }

  if (!options.includeBody) {
    return `${entries.map(formatListRow).join("\n")}\n`;
  }

  return `${entries.map((entry) => `${formatListRow(entry)}\n${resolveBody(entry)}\n`).join("\n")}`;
}

export function formatEmail(entry: NormalizedMailboxMessage): string {
  const { message } = entry;
  const lines = [
    `Mailbox: ${entry.mailbox}`,
    `Id:      ${message.id ?? ""}`,
    `From:    ${message.from || "Unknown sender"}`,
    message.to.length > 0 ? `To:      ${message.to.join(", ")}` : null,
    message.cc.length > 0 ? `Cc:      ${message.cc.join(", ")}` : null,
    `Received: ${message.receivedAt}`,
    `Subject: ${message.subject || "(no subject)"}`,
    "",
    resolveBody(entry),
  ].filter((line): line is string => line !== null);

  return `${lines.join("\n")}\n`;
}

function formatListRow(entry: NormalizedMailboxMessage): string {
  const { message } = entry;
  return [
    entry.mailbox,
    message.receivedAt || "(no date)",
    message.id ?? "",
    message.subject || "(no subject)",
    message.from || "(unknown)",
  ].join(SEP);
}

function resolveBody(entry: NormalizedMailboxMessage): string {
  const bodyText = entry.message.bodyText.trim();
  if (bodyText) {
    return bodyText;
  }
  const preview = entry.message.preview.trim();
  return preview || "(no body)";
}
