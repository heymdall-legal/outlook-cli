import { mkdir, readFile, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

const KEYRING_SERVICE_NAME = "outlook";
const KEYRING_ACCOUNT_NAME = "cli";

export function getConfigPath(home = homedir()): string {
  return join(home, ".config", "outlook");
}

export function getFileConfigPath(home = homedir()): string {
  return join(getConfigPath(home), "config.json");
}

export type FileConfig = {
  email?: string;
  username?: string;
  host?: string;
  endpoint?: string;
  protocolVersion?: string;
};

export type ResolvedConfig = {
  email: string;
  username: string;
  host: string;
  password: string;
  endpoint?: string;
  protocolVersion?: string;
};

export async function loadConfig(): Promise<ResolvedConfig> {
  const { email, username, host, endpoint, protocolVersion } = await readConfigFile(getFileConfigPath());
  const password = await readPassword();

  if (!email || !username || !host || !password) {
    throw new Error("Missing required configuration: email, username, host or password. Run `outlook config init`.");
  }

  return { email, username, host, password, endpoint, protocolVersion };
}

async function readConfigFile(filePath: string): Promise<FileConfig> {
  try {
    return JSON.parse(await readFile(filePath, "utf8")) as FileConfig;
  } catch (error) {
    if (isMissingFile(error)) {
      return {};
    }
    throw error;
  }
}

export async function writeConfig(config: ResolvedConfig): Promise<void> {
  const configPath = getFileConfigPath();
  const { password, ...fileConfig } = config;

  await mkdir(dirname(configPath), { recursive: true });
  await writeFile(configPath, `${JSON.stringify(fileConfig, null, 2)}\n`, "utf8");
  await savePassword(password);
}

async function getKeytar() {
  try {
    const { default: keytar } = await import("keytar");
    return keytar;
  } catch {
    throw new Error(
      "keytar is not available. Install libsecret (libsecret-1-dev on Debian/Ubuntu, libsecret-devel on RHEL/Fedora) and reinstall.",
    );
  }
}

async function readPassword() {
  const keytar = await getKeytar();
  const results = await keytar.findCredentials(KEYRING_SERVICE_NAME);
  if (!results || results.length === 0) {
    return null;
  }
  return results.find((cred) => cred.account === KEYRING_ACCOUNT_NAME)?.password ?? null;
}

export async function savePassword(password: string) {
  const keytar = await getKeytar();
  await keytar.setPassword(KEYRING_SERVICE_NAME, KEYRING_ACCOUNT_NAME, password);
}

export function redactConfig(config: ResolvedConfig): ResolvedConfig {
  return { ...config, password: "****" };
}

function isMissingFile(error: unknown): error is NodeJS.ErrnoException {
  return typeof error === "object" && error !== null && "code" in error && error.code === "ENOENT";
}
