import type { ActiveSyncClientConfig } from "../types.js";

export interface DiscoveryAttempt {
  endpoint: string;
  status?: number;
  code?: string;
  message?: string;
}

export interface DiscoveryResult {
  endpoint: string;
  attempts: DiscoveryAttempt[];
}

interface FetchResponseLike {
  status: number;
}

type FetchLike = (
  input: string,
  init: {
    method: string;
    redirect: "manual";
    headers: Record<string, string>;
  },
) => Promise<FetchResponseLike>;

interface HttpErrorInfo {
  status?: number;
  code?: string;
}

interface DiscoveryConfig
  extends Pick<
    ActiveSyncClientConfig,
    "endpoint" | "host" | "email" | "userAgent" | "protocolVersion" | "username" | "password"
  > {}

export function getDefaultEndpoint(host: string): string {
  return `https://${host}/Microsoft-Server-ActiveSync`;
}

export function createDiscoveryCandidates({
  endpoint,
  host,
  email,
}: Pick<DiscoveryConfig, "endpoint" | "host" | "email">): string[] {
  if (endpoint) {
    return [endpoint];
  }

  const candidates: string[] = [];
  const domain = host || email.split("@")[1];

  if (!domain) {
    return candidates;
  }

  candidates.push(getDefaultEndpoint(domain));

  if (!domain.startsWith("autodiscover.")) {
    candidates.push(getDefaultEndpoint(`autodiscover.${domain}`));
  }

  if (!domain.startsWith("mail.")) {
    candidates.push(getDefaultEndpoint(`mail.${domain}`));
  }

  return [...new Set(candidates)];
}

export function classifyHttpError({ status, code }: HttpErrorInfo = {}): string {
  if (code) {
    return "network-failure";
  }

  if (status === 401) {
    return "bad-credentials";
  }

  if (status === 403) {
    return "device-blocked";
  }

  if (status === 404) {
    return "endpoint-not-found";
  }

  if (status !== undefined && status >= 400) {
    return "protocol-error";
  }

  return "unknown";
}

export async function discoverEndpoint(
  config: DiscoveryConfig,
  request: FetchLike = fetch as FetchLike,
): Promise<DiscoveryResult> {
  const candidates = createDiscoveryCandidates(config);
  const attempts: DiscoveryAttempt[] = [];

  for (const candidate of candidates) {
    try {
      const response = await request(candidate, {
        method: "OPTIONS",
        redirect: "manual",
        headers: buildDiscoveryHeaders(config),
      });

      attempts.push({ endpoint: candidate, status: response.status });

      if ([200, 401, 403, 451].includes(response.status)) {
        return { endpoint: candidate, attempts };
      }
    } catch (error) {
      const failure = error as {
        code?: string;
        cause?: { code?: string };
        message?: string;
      };
      attempts.push({
        endpoint: candidate,
        code: failure.code ?? failure.cause?.code ?? "network-error",
        message: failure.message ?? "Unknown error",
      });
    }
  }

  throw new Error(`Unable to discover a working ActiveSync endpoint. Attempts: ${JSON.stringify(attempts)}`);
}

function buildDiscoveryHeaders(config: DiscoveryConfig): Record<string, string> {
  return {
    "User-Agent": config.userAgent,
    "MS-ASProtocolVersion": config.protocolVersion,
    Authorization: `Basic ${Buffer.from(`${config.username}:${config.password}`).toString("base64")}`,
  };
}
