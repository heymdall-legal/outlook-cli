import { z } from "zod";
import type { Plugin } from "../types.js";
import { getOutlookEmails } from "./client.js";
import { mapMailboxMessageToMemoryDraft } from "./mapping.js";
import { filterMessagesSince } from "./messages.js";

export const outlookConfigSchema = z.object({
  email: z.string().min(1),
  username: z.string().min(1),
  password: z.string().min(1),
  host: z.string().min(1),
  initialLookbackDays: z.coerce.number().int().min(1).default(1),
});

export type OutlookPluginConfig = z.infer<typeof outlookConfigSchema>;

export const outlookPlugin: Plugin<OutlookPluginConfig> = {
  id: "outlook",
  name: "Outlook",
  defaultSchedule: "45 23 * * *",
  producedMemoryTypes: ["outlook.email"],
  configSchema: outlookConfigSchema,
  async run({ config, logger, lastRunAt }) {
    logger.info(
      {
        lastRunAt,
        initialLookbackDays: config.initialLookbackDays,
      },
      "outlook: fetching mailbox messages",
    );

    const messages = await getOutlookEmails({
      email: config.email,
      username: config.username,
      password: config.password,
      host: config.host,
      days: config.initialLookbackDays,
    });

    const filtered = lastRunAt ? filterMessagesSince(messages, lastRunAt) : messages;

    return filtered.flatMap((entry) => {
      const draft = mapMailboxMessageToMemoryDraft(entry);
      return draft ? [draft] : [];
    });
  },
};
