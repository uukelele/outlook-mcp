export type AuthConfig =
  | { kind: "token"; token: string }
  | { kind: "token-file"; path: string }
  | { kind: "command"; command: string }
  | { kind: "azure-cli"; resource: string; tenant?: string }
  | { kind: "cookie"; cookie: string; canary?: string };

export interface Config {
  /** Origin of Outlook on the web, e.g. https://outlook.cloud.microsoft */
  origin: string;
  /** Origin that serves attachment downloads. */
  attachmentsOrigin: string;
  auth: AuthConfig;
  /** Explicit X-AnchorMailbox value; derived from the token when unset. */
  anchorMailbox?: string;
  /** Windows time zone ID used to interpret times without an offset. Defaults to the mailbox setting. */
  timeZone?: string;
  userAgent: string;
  requestTimeoutMs: number;
  maxDownloadBytes: number;
  readOnly: boolean;
  debug: boolean;
}

const AUTH_KINDS = ["auto", "token", "token-file", "command", "azure-cli", "cookie"] as const;

function env(name: string, source: NodeJS.ProcessEnv): string | undefined {
  const value = source[name]?.trim();
  return value ? value : undefined;
}

function flag(name: string, source: NodeJS.ProcessEnv): boolean {
  return /^(1|true|yes|on)$/i.test(source[name]?.trim() ?? "");
}

function positiveNumber(name: string, fallback: number, source: NodeJS.ProcessEnv): number {
  const raw = env(name, source);
  if (raw === undefined) return fallback;
  const value = Number(raw);
  if (!Number.isFinite(value) || value <= 0) {
    throw new Error(`${name} must be a positive number, got "${raw}"`);
  }
  return value;
}

function stripTrailingSlash(url: string): string {
  return url.replace(/\/+$/, "");
}

function resolveAuth(source: NodeJS.ProcessEnv): AuthConfig {
  const mode = (env("OUTLOOK_AUTH", source) ?? "auto").toLowerCase();
  if (!(AUTH_KINDS as readonly string[]).includes(mode)) {
    throw new Error(`OUTLOOK_AUTH must be one of ${AUTH_KINDS.join(", ")}; got "${mode}"`);
  }
  const token = env("OUTLOOK_ACCESS_TOKEN", source);
  const tokenFile = env("OUTLOOK_TOKEN_FILE", source);
  const command = env("OUTLOOK_TOKEN_COMMAND", source);
  const cookie = env("OUTLOOK_COOKIE", source);
  const azure = (): AuthConfig => ({
    kind: "azure-cli",
    resource: env("OUTLOOK_AZ_RESOURCE", source) ?? "https://outlook.office.com",
    tenant: env("OUTLOOK_AZ_TENANT", source),
  });
  const require = <T>(value: T | undefined, variable: string): T => {
    if (value === undefined) throw new Error(`OUTLOOK_AUTH=${mode} requires ${variable} to be set`);
    return value;
  };

  switch (mode) {
    case "token":
      return { kind: "token", token: require(token, "OUTLOOK_ACCESS_TOKEN") };
    case "token-file":
      return { kind: "token-file", path: require(tokenFile, "OUTLOOK_TOKEN_FILE") };
    case "command":
      return { kind: "command", command: require(command, "OUTLOOK_TOKEN_COMMAND") };
    case "cookie":
      return { kind: "cookie", cookie: require(cookie, "OUTLOOK_COOKIE"), canary: env("OUTLOOK_CANARY", source) };
    case "azure-cli":
      return azure();
    default:
      if (token) return { kind: "token", token };
      if (tokenFile) return { kind: "token-file", path: tokenFile };
      if (command) return { kind: "command", command };
      if (cookie) return { kind: "cookie", cookie, canary: env("OUTLOOK_CANARY", source) };
      return azure();
  }
}

export function loadConfig(source: NodeJS.ProcessEnv = process.env): Config {
  return {
    origin: stripTrailingSlash(env("OUTLOOK_ORIGIN", source) ?? "https://outlook.cloud.microsoft"),
    attachmentsOrigin: stripTrailingSlash(env("OUTLOOK_ATTACHMENTS_ORIGIN", source) ?? "https://attachments.office.net"),
    auth: resolveAuth(source),
    anchorMailbox: env("OUTLOOK_ANCHOR_MAILBOX", source),
    timeZone: env("OUTLOOK_TIMEZONE", source),
    userAgent:
      env("OUTLOOK_USER_AGENT", source) ??
      "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
    requestTimeoutMs: positiveNumber("OUTLOOK_REQUEST_TIMEOUT_MS", 30_000, source),
    maxDownloadBytes: Math.floor(positiveNumber("OUTLOOK_MAX_DOWNLOAD_MB", 25, source) * 1024 * 1024),
    readOnly: flag("OUTLOOK_READ_ONLY", source),
    debug: flag("OUTLOOK_DEBUG", source),
  };
}
