import readline from "node:readline";
import { Writable } from "node:stream";

export type OutputFormat = "text" | "json";

type PrintOutputParams = {
  value: unknown;
  format?: OutputFormat;
  textFormatter?: (value: unknown) => string;
};

export function printOutput({ value, format = "text", textFormatter }: PrintOutputParams): void {
  if (format === "json") {
    process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
    return;
  }

  if (textFormatter) {
    process.stdout.write(textFormatter(value));
    return;
  }

  process.stdout.write(`${formatLine(value)}\n`);
}

export function formatLine(value: unknown): string {
  return typeof value === "string" ? value : JSON.stringify(value, null, 2);
}

export function parseOutputFormat(value: string | undefined): OutputFormat {
  if (value === undefined) {
    return "text";
  }
  if (value !== "text" && value !== "json") {
    throw new Error(`Unsupported output format: ${value}. Use "text" or "json".`);
  }
  return value;
}

export function formatError(error: unknown): string {
  if (!(error instanceof Error)) {
    return `${String(error)}\n`;
  }

  type ErrorWithDetails = Error & { cause?: unknown; code?: unknown };
  const messages = [error.message];
  let { cause } = error as ErrorWithDetails;

  while (cause) {
    if (cause instanceof Error) {
      const code = typeof (cause as ErrorWithDetails).code === "string" ? ` (${(cause as ErrorWithDetails).code})` : "";
      messages.push(`${cause.message}${code}`);
      cause = (cause as ErrorWithDetails).cause;
    } else {
      messages.push(String(cause));
      cause = undefined;
    }
  }

  return `${messages.join("\nCaused by: ")}\n`;
}

export class MutableStdout extends Writable {
  public muted = false;

  // eslint-disable-next-line no-underscore-dangle, class-methods-use-this
  _write(chunk: Buffer, _encoding: unknown, callback: (error?: Error | null) => void) {
    if (this.muted) {
      const str = chunk.toString();
      // eslint-disable-next-line no-control-regex
      if (/^(?:\x1B|\r|\n|\x08|\x7f)/.test(str)) {
        process.stdout.write(chunk);
      } else {
        process.stdout.write("*".repeat(str.length));
      }
    } else {
      process.stdout.write(chunk);
    }
    callback();
  }
}

type PromptParams = {
  message: string;
  defaultValue?: string;
  secret?: boolean;
};

type PromptDependencies = {
  createInterface?: typeof readline.createInterface;
};

export async function prompt(
  { message, defaultValue = "", secret = false }: PromptParams,
  dependencies: PromptDependencies = {},
): Promise<string> {
  const outputStream = new MutableStdout();
  const createInterface = dependencies.createInterface ?? readline.createInterface;
  const rl = createInterface({ input: process.stdin, output: outputStream, terminal: true });

  try {
    return await new Promise<string>((resolve) => {
      rl.question(
        `${message.trim()}${defaultValue ? ` (leave empty for default, ${defaultValue})` : ""}: `,
        (answer: string) => {
          outputStream.muted = false;
          if (secret) {
            process.stdout.write("\n");
          }
          resolve(answer || defaultValue);
        },
      );
      if (secret) {
        outputStream.muted = true;
      }
    });
  } finally {
    rl.close();
  }
}
