import { exec, execFile } from "node:child_process";
import { readFile } from "node:fs/promises";
import { promisify } from "node:util";
import type { AuthConfig } from "../config.js";
import { AuthError } from "../errors.js";
import { checkAudience, decodeJwt, userPrincipal, type TokenClaims } from "./jwt.js";

const execFileAsync = promisify(execFile);
const execAsync = promisify(exec);

export interface Identity {
  upn?: string;
  name?: string;
  puid?: string;
  tenantId?: string;
}

export interface AuthProvider {
  readonly description: string;
  /** Whether a failed request can be retried with freshly obtained credentials. */
  readonly canRefresh: boolean;
  headers(options?: { forceRefresh?: boolean }): Promise<Record<string, string>>;
  /** Identity derived from the credentials, if they carry one (JWT claims). */
  identity(): Identity | undefined;
}

export interface TokenResult {
  token: string;
  /** Expiry in epoch milliseconds, if known independently of the JWT. */
  expiresAt?: number;
}

export interface TokenSource {
  readonly description: string;
  readonly refreshable: boolean;
  fetch(): Promise<TokenResult>;
}

type Logger = (message: string) => void;

/** Refresh this long before the token actually expires. */
const EXPIRY_SKEW_MS = 5 * 60 * 1000;
/** How long to trust a token whose expiry cannot be determined. */
const UNKNOWN_EXPIRY_TTL_MS = 10 * 60 * 1000;

export class BearerAuthProvider implements AuthProvider {
  private cached?: { token: string; expiresAt: number; claims?: TokenClaims };
  private inflight?: Promise<string>;
  private warnedAudience = false;

  constructor(
    private readonly source: TokenSource,
    private readonly log: Logger = () => {},
    private readonly now: () => number = Date.now,
  ) {}

  get description(): string {
    return this.source.description;
  }

  get canRefresh(): boolean {
    return this.source.refreshable;
  }

  async headers(options: { forceRefresh?: boolean } = {}): Promise<Record<string, string>> {
    return { Authorization: `Bearer ${await this.token(options.forceRefresh)}` };
  }

  identity(): Identity | undefined {
    const claims = this.cached?.claims;
    if (!claims) return undefined;
    return {
      upn: userPrincipal(claims),
      name: typeof claims.name === "string" ? claims.name : undefined,
      puid: typeof claims.puid === "string" ? claims.puid : undefined,
      tenantId: typeof claims.tid === "string" ? claims.tid : undefined,
    };
  }

  async token(forceRefresh = false): Promise<string> {
    if (!forceRefresh && this.cached && this.cached.expiresAt - EXPIRY_SKEW_MS > this.now()) {
      return this.cached.token;
    }
    if (forceRefresh && this.cached && !this.source.refreshable) {
      throw new AuthError(
        `The Outlook access token from ${this.source.description} was rejected and cannot be refreshed automatically. ` +
          "Provide a new token (or switch to Azure CLI auth) and restart the server.",
      );
    }
    this.inflight ??= this.load().finally(() => {
      this.inflight = undefined;
    });
    return this.inflight;
  }

  private async load(): Promise<string> {
    const { token, expiresAt } = await this.source.fetch();
    if (!token) throw new AuthError(`${this.source.description} returned an empty token`);
    const claims = decodeJwt(token);
    const audience = checkAudience(claims);
    if (audience === "graph") {
      throw new AuthError(
        "The access token is for Microsoft Graph (aud=graph.microsoft.com). Outlook on the web needs a token for " +
          "https://outlook.office.com, e.g. `az account get-access-token --resource https://outlook.office.com`.",
      );
    }
    if (audience === "unknown" && claims && !this.warnedAudience) {
      this.warnedAudience = true;
      this.log(`warning: token audience "${String(claims.aud)}" is not a known Outlook audience; requests may fail`);
    }
    const jwtExpiry = typeof claims?.exp === "number" ? claims.exp * 1000 : undefined;
    this.cached = {
      token,
      claims,
      expiresAt: expiresAt ?? jwtExpiry ?? this.now() + UNKNOWN_EXPIRY_TTL_MS,
    };
    return token;
  }
}

export class StaticTokenSource implements TokenSource {
  readonly description = "OUTLOOK_ACCESS_TOKEN";
  readonly refreshable = false;
  constructor(private readonly token: string) {}
  async fetch(): Promise<TokenResult> {
    return { token: stripBearer(this.token) };
  }
}

export class FileTokenSource implements TokenSource {
  readonly refreshable = true;
  constructor(private readonly path: string) {}
  get description(): string {
    return `token file ${this.path}`;
  }
  async fetch(): Promise<TokenResult> {
    let text: string;
    try {
      text = await readFile(this.path, "utf8");
    } catch (error) {
      throw new AuthError(`Cannot read OUTLOOK_TOKEN_FILE (${this.path}): ${(error as Error).message}`);
    }
    return parseTokenOutput(text, this.description);
  }
}

export class CommandTokenSource implements TokenSource {
  readonly refreshable = true;
  constructor(private readonly command: string) {}
  get description(): string {
    return "OUTLOOK_TOKEN_COMMAND";
  }
  async fetch(): Promise<TokenResult> {
    try {
      const { stdout } = await execAsync(this.command, { timeout: 60_000, maxBuffer: 1024 * 1024 });
      return parseTokenOutput(stdout, this.description);
    } catch (error) {
      if (error instanceof AuthError) throw error;
      const stderr = String((error as { stderr?: string }).stderr ?? "").trim();
      throw new AuthError(`OUTLOOK_TOKEN_COMMAND failed: ${stderr || (error as Error).message}`);
    }
  }
}

export class AzureCliTokenSource implements TokenSource {
  readonly refreshable = true;
  constructor(
    private readonly resource: string,
    private readonly tenant?: string,
  ) {}

  get description(): string {
    return `Azure CLI (resource ${this.resource})`;
  }

  async fetch(): Promise<TokenResult> {
    // On Windows the CLI is a .cmd, so execFile must go through the shell; the
    // shell joins arguments without quoting, so validate anything interpolated.
    validateShellWord(this.resource, "OUTLOOK_AZ_RESOURCE", /^[A-Za-z0-9:/.-]+$/);
    if (this.tenant) validateShellWord(this.tenant, "OUTLOOK_AZ_TENANT", /^[A-Za-z0-9.-]+$/);
    const args = ["account", "get-access-token", "--resource", this.resource, "--output", "json"];
    if (this.tenant) args.push("--tenant", this.tenant);
    let stdout: string;
    try {
      ({ stdout } = await execFileAsync("az", args, {
        timeout: 60_000,
        maxBuffer: 1024 * 1024,
        shell: process.platform === "win32",
      }));
    } catch (error) {
      const err = error as NodeJS.ErrnoException & { stderr?: string };
      if (err.code === "ENOENT") {
        throw new AuthError(
          "Azure CLI (`az`) was not found. Install it and run `az login`, or set OUTLOOK_ACCESS_TOKEN / OUTLOOK_TOKEN_COMMAND.",
        );
      }
      const stderr = String(err.stderr ?? "").trim();
      const hint = /az login|AADSTS|expired|interaction/i.test(stderr) ? " Run `az login` and try again." : "";
      throw new AuthError(`Azure CLI could not issue an Outlook token: ${lastLine(stderr) || err.message}.${hint}`);
    }
    let parsed: { accessToken?: string; expires_on?: number | string; expiresOn?: string };
    try {
      parsed = JSON.parse(stdout);
    } catch {
      throw new AuthError("Azure CLI returned output that is not JSON");
    }
    if (!parsed.accessToken) throw new AuthError("Azure CLI did not return an accessToken");
    const epoch = Number(parsed.expires_on);
    return {
      token: parsed.accessToken,
      expiresAt: Number.isFinite(epoch) && epoch > 0 ? epoch * 1000 : undefined,
    };
  }
}

/** Cookie + canary session (e.g. on-premises OWA). Cannot be refreshed automatically. */
export class CookieAuthProvider implements AuthProvider {
  readonly description = "OWA session cookie";
  readonly canRefresh = false;
  private readonly canary?: string;

  constructor(
    private readonly cookie: string,
    canary?: string,
  ) {
    this.canary = canary ?? /(?:^|;\s*)X-OWA-CANARY=([^;]+)/i.exec(cookie)?.[1];
  }

  async headers(options: { forceRefresh?: boolean } = {}): Promise<Record<string, string>> {
    if (options.forceRefresh) {
      throw new AuthError("The OWA session cookie was rejected (expired?). Capture a fresh OUTLOOK_COOKIE and restart.");
    }
    const headers: Record<string, string> = { Cookie: this.cookie };
    if (this.canary) headers["X-OWA-CANARY"] = this.canary;
    return headers;
  }

  identity(): Identity | undefined {
    return undefined;
  }
}

export function createAuthProvider(config: AuthConfig, log: Logger): AuthProvider {
  switch (config.kind) {
    case "token":
      return new BearerAuthProvider(new StaticTokenSource(config.token), log);
    case "token-file":
      return new BearerAuthProvider(new FileTokenSource(config.path), log);
    case "command":
      return new BearerAuthProvider(new CommandTokenSource(config.command), log);
    case "azure-cli":
      return new BearerAuthProvider(new AzureCliTokenSource(config.resource, config.tenant), log);
    case "cookie":
      return new CookieAuthProvider(config.cookie, config.canary);
  }
}

/** Accepts a bare token, "Bearer <token>", or JSON containing accessToken/access_token. */
export function parseTokenOutput(output: string, description: string): TokenResult {
  const text = output.trim();
  if (!text) throw new AuthError(`${description} produced no token`);
  if (text.startsWith("{")) {
    let json: Record<string, unknown>;
    try {
      json = JSON.parse(text);
    } catch {
      throw new AuthError(`${description} produced invalid JSON`);
    }
    const token = json.accessToken ?? json.access_token ?? json.token;
    if (typeof token !== "string" || !token) {
      throw new AuthError(`${description} JSON has no accessToken/access_token field`);
    }
    const epoch = Number(json.expires_on ?? json.expiresOn);
    return { token: stripBearer(token), expiresAt: Number.isFinite(epoch) && epoch > 0 ? epoch * 1000 : undefined };
  }
  return { token: stripBearer(text.split(/\s*\n/)[0]!) };
}

function stripBearer(token: string): string {
  return token.trim().replace(/^Bearer\s+/i, "");
}

/** Reject values that could smuggle shell syntax when interpolated on Windows. */
function validateShellWord(value: string, name: string, pattern: RegExp): void {
  if (!pattern.test(value)) {
    throw new AuthError(`${name} contains unsupported characters: "${value}" (expected e.g. https://outlook.office.com)`);
  }
}

function lastLine(text: string): string {
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  return lines[lines.length - 1] ?? "";
}
