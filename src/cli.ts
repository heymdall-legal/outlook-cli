import { Command } from "commander";

import { OutlookClient } from "./client.js";
import { getFileConfigPath, loadConfig, redactConfig, writeConfig, type ResolvedConfig } from "./config.js";
import { formatEmail, formatEmailList } from "./format.js";
import { formatError, parseOutputFormat, printOutput, prompt } from "./io.js";
import type { MailboxKind, NormalizedMailboxMessage } from "./types.js";

function parseMailbox(value: string | undefined): MailboxKind | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (value !== "inbox" && value !== "sent") {
    throw new Error(`Unsupported mailbox: ${value}. Use "inbox" or "sent".`);
  }
  return value;
}

function parsePositiveInt(value: string | undefined, fallback: number, label: string): number {
  if (value === undefined) {
    return fallback;
  }
  const parsed = Number.parseInt(value, 10);
  if (Number.isNaN(parsed) || parsed <= 0) {
    throw new Error(`${label} must be a positive integer, got "${value}".`);
  }
  return parsed;
}

export function createCli() {
  const program = new Command();
  program.name("outlook").description("Read Outlook/Exchange mail over ActiveSync from the command line.");

  const configCommand = program.command("config").description("Create, verify, and inspect the connection settings.");

  configCommand
    .command("init")
    .description("Prompt for connection settings, verify them, and save them locally.")
    .action(async () => {
      const email = await prompt({ message: "Email" });
      const username = await prompt({ message: "Username", defaultValue: email });
      const host = await prompt({ message: "Host (e.g. mail.example.com)" });
      const password = await prompt({ message: "Password", secret: true });

      const config: ResolvedConfig = { email, username, host, password };
      await new OutlookClient(config).verifyConnection();
      await writeConfig(config);

      process.stdout.write(`Verified and saved config to ${getFileConfigPath()}\n`);
    });

  configCommand
    .command("verify")
    .description("Verify the saved configuration by connecting to the server.")
    .action(async () => {
      await new OutlookClient(await loadConfig()).verifyConnection();
      process.stdout.write("Connection verified.\n");
    });

  configCommand
    .command("show")
    .description("Print the saved configuration with the password redacted.")
    .action(async () => {
      printOutput({ value: redactConfig(await loadConfig()), format: "json" });
    });

  const emailsCommand = program.command("emails").description("List and fetch emails.");

  emailsCommand
    .command("list")
    .description("List emails received in the last N hours.")
    .option("--hours <N>", "Lookback window in hours.", "24")
    .option("--mailbox <inbox|sent>", "Restrict to one mailbox. Defaults to both.")
    .option("--body", "Include the full plain-text body for each email.")
    .option("--max-pages <N>", "Safety cap on sync pages per mailbox.", "50")
    .option("--output <text|json>", "Output format: text or json.", "text")
    .action(
      async (options: { hours?: string; mailbox?: string; body?: boolean; maxPages?: string; output?: string }) => {
        const format = parseOutputFormat(options.output);
        const hours = parsePositiveInt(options.hours, 24, "--hours");
        const maxPages = parsePositiveInt(options.maxPages, 50, "--max-pages");
        const mailbox = parseMailbox(options.mailbox);

        const client = new OutlookClient(await loadConfig());
        const emails = await client.listEmails({ hours, mailbox, maxPages });

        printOutput({
          value: emails,
          format,
          textFormatter: (value) =>
            formatEmailList(value as NormalizedMailboxMessage[], { includeBody: Boolean(options.body) }),
        });
      },
    );

  emailsCommand
    .command("get <id>")
    .description("Fetch a single full email by id.")
    .option("--output <text|json>", "Output format: text or json.", "text")
    .action(async (id: string, options: { output?: string }) => {
      const format = parseOutputFormat(options.output);
      const client = new OutlookClient(await loadConfig());
      const email = await client.getEmail({ id });

      if (!email) {
        throw new Error(`Email with id "${id}" was not found in the Inbox or Sent folder.`);
      }

      printOutput({
        value: email,
        format,
        textFormatter: (value) => formatEmail(value as NormalizedMailboxMessage),
      });
    });

  return {
    program,
    async run(argv: string[]): Promise<number> {
      try {
        await program.parseAsync(argv, { from: "user" });
        return 0;
      } catch (error) {
        process.stderr.write(formatError(error));
        return 1;
      }
    },
  };
}
