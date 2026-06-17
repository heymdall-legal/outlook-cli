import type { NormalizedMailboxMessage } from "./types.js";

export function filterRecentMessages(
  messages: NormalizedMailboxMessage[],
  days: number,
  now: Date = new Date(),
): NormalizedMailboxMessage[] {
  const cutoff = new Date(now);
  cutoff.setUTCDate(cutoff.getUTCDate() - days);

  return messages.filter((entry) => {
    const receivedAt = new Date(entry.message.receivedAt);
    return !Number.isNaN(receivedAt.valueOf()) && receivedAt >= cutoff;
  });
}

export function filterMessagesSince(
  messages: NormalizedMailboxMessage[],
  lastRunAt: string,
): NormalizedMailboxMessage[] {
  const cutoff = new Date(lastRunAt);

  return messages.filter((entry) => {
    const receivedAt = new Date(entry.message.receivedAt);
    return !Number.isNaN(receivedAt.valueOf()) && receivedAt >= cutoff;
  });
}

export function filterMessagesWithinHours(
  messages: NormalizedMailboxMessage[],
  hours: number,
  now: Date = new Date(),
): NormalizedMailboxMessage[] {
  const cutoff = new Date(now.getTime() - hours * 60 * 60 * 1000);

  return messages.filter((entry) => {
    const receivedAt = new Date(entry.message.receivedAt);
    return !Number.isNaN(receivedAt.valueOf()) && receivedAt >= cutoff;
  });
}
