import { randomUUID } from "node:crypto";
import type { AuthProvider, Identity } from "../auth/providers.js";
import { AuthError, OwaHttpError, OwaResponseError } from "../errors.js";
import { envelope, type EnvelopeOptions, type Json } from "./builders.js";

export type FetchLike = (input: string, init: RequestInit) => Promise<Response>;

export interface OwaClientOptions {
  origin: string;
  attachmentsOrigin: string;
  auth: AuthProvider;
  anchorMailbox?: string;
  userAgent: string;
  requestTimeoutMs: number;
  fetch?: FetchLike;
  log?: (message: string) => void;
  debug?: boolean;
  /** Sleep used for retry back-off; injectable for tests. */
  sleep?: (ms: number) => Promise<void>;
}

export interface CallOptions extends EnvelopeOptions {
  /** Routing hint sent as `app=` and X-Req-Source. */
  app?: "Mail" | "Calendar";
  /** Send the payload as-is instead of wrapping it in the standard envelope. */
  raw?: boolean;
}

const MAX_THROTTLE_RETRIES = 2;
const MAX_RETRY_AFTER_MS = 15_000;

export class OwaClient {
  private sequence = 0;
  private readonly sessionId = randomUUID();
  private readonly fetchImpl: FetchLike;
  private readonly sleep: (ms: number) => Promise<void>;
  private anchor?: string;

  constructor(private readonly options: OwaClientOptions) {
    this.fetchImpl = options.fetch ?? ((input, init) => fetch(input, init));
    this.sleep = options.sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
    this.anchor = options.anchorMailbox;
  }

  get origin(): string {
    return this.options.origin;
  }

  get attachmentsOrigin(): string {
    return this.options.attachmentsOrigin;
  }

  /** Identity from the current credentials; populated after the first authenticated request. */
  identity(): Identity | undefined {
    return this.options.auth.identity();
  }

  /** Call a service.svc action and return the parsed JSON response. Does not inspect ResponseClass. */
  async call<T = any>(action: string, body: Json, options: CallOptions = {}): Promise<T> {
    const payload = options.raw ? body : envelope(action, body, options);
    const app = options.app ?? "Mail";
    let forceRefresh = false;
    let throttleRetries = 0;

    for (;;) {
      const authHeaders = await this.options.auth.headers({ forceRefresh });
      const url = `${this.options.origin}/owa/service.svc?action=${encodeURIComponent(action)}&app=${app}&n=${++this.sequence}`;
      const headers: Record<string, string> = {
        ...authHeaders,
        Action: action,
        "Content-Type": "application/json; charset=utf-8",
        Accept: "application/json",
        Prefer: 'IdType="ImmutableId", exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"',
        "X-Requested-With": "XMLHttpRequest",
        "X-OWA-CorrelationId": randomUUID(),
        "X-OWA-SessionId": this.sessionId,
        "X-Req-Source": app,
        "client-request-id": randomUUID(),
        Origin: this.options.origin,
        Referer: `${this.options.origin}/${app === "Calendar" ? "calendar" : "mail"}/`,
        "User-Agent": this.options.userAgent,
      };
      const anchor = this.anchorMailbox();
      if (anchor) headers["X-AnchorMailbox"] = anchor;

      this.debug(`→ ${action} #${this.sequence}`);
      let response: Response;
      try {
        response = await this.fetchImpl(url, {
          method: "POST",
          headers,
          body: JSON.stringify(payload),
          redirect: "manual",
          signal: AbortSignal.timeout(this.options.requestTimeoutMs),
        });
      } catch (error) {
        const reason = (error as Error).name === "TimeoutError" ? `timed out after ${this.options.requestTimeoutMs} ms` : (error as Error).message;
        throw new OwaHttpError(`${action} request failed: ${reason}`, 0, action);
      }
      this.debug(`← ${action} #${this.sequence} HTTP ${response.status}`);

      if (isAuthFailure(response)) {
        await drain(response);
        if (!forceRefresh && this.options.auth.canRefresh) {
          forceRefresh = true;
          continue;
        }
        throw new AuthError(
          `Outlook rejected the credentials from ${this.options.auth.description} (HTTP ${response.status}). ` +
            "The token may be expired or lack mailbox access.",
        );
      }

      if ((response.status === 429 || response.status === 503) && throttleRetries < MAX_THROTTLE_RETRIES) {
        throttleRetries++;
        await drain(response);
        await this.sleep(retryDelay(response, throttleRetries));
        continue;
      }

      return parseResponse<T>(action, response);
    }
  }

  /**
   * Call an action whose response uses Body.ResponseMessages and return the
   * response message items. Throws on the first item with ResponseClass "Error".
   */
  async callForMessages(action: string, body: Json, options: CallOptions = {}): Promise<Json[]> {
    const result = await this.call(action, body, options);
    const messages = responseMessages(result);
    for (const message of messages) assertSuccess(action, message);
    return messages;
  }

  /** GET a binary resource (attachment download). */
  async download(url: string, headers: Record<string, string>, maxBytes: number): Promise<{ data: Buffer; contentType?: string; fileName?: string }> {
    let response: Response;
    try {
      response = await this.fetchImpl(url, {
        method: "GET",
        headers: { Origin: this.options.origin, Accept: "*/*", "User-Agent": this.options.userAgent, ...headers },
        redirect: "manual",
        signal: AbortSignal.timeout(Math.max(this.options.requestTimeoutMs, 120_000)),
      });
    } catch (error) {
      throw new OwaHttpError(`Attachment download failed: ${(error as Error).message}`, 0, "GetFileAttachment");
    }
    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new OwaHttpError(`Attachment download failed with HTTP ${response.status}${text ? `: ${truncate(text, 200)}` : ""}`, response.status, "GetFileAttachment");
    }
    const declared = Number(response.headers.get("content-length"));
    if (Number.isFinite(declared) && declared > maxBytes) {
      await drain(response);
      throw new OwaHttpError(`Attachment is ${declared} bytes, above the ${maxBytes}-byte limit (OUTLOOK_MAX_DOWNLOAD_MB)`, 413, "GetFileAttachment");
    }
    const data = Buffer.from(await response.arrayBuffer());
    if (data.length > maxBytes) {
      throw new OwaHttpError(`Attachment is ${data.length} bytes, above the ${maxBytes}-byte limit (OUTLOOK_MAX_DOWNLOAD_MB)`, 413, "GetFileAttachment");
    }
    return {
      data,
      contentType: response.headers.get("content-type") ?? undefined,
      fileName: fileNameFromDisposition(response.headers.get("content-disposition")),
    };
  }

  private anchorMailbox(): string | undefined {
    if (this.anchor) return this.anchor;
    const identity = this.options.auth.identity();
    if (identity?.puid && identity.tenantId) {
      this.anchor = `PUID:${identity.puid}@${identity.tenantId}`;
    } else if (identity?.upn) {
      this.anchor = identity.upn;
    }
    return this.anchor;
  }

  private debug(message: string): void {
    if (this.options.debug) this.options.log?.(message);
  }
}

function isAuthFailure(response: Response): boolean {
  if (response.status === 401 || response.status === 440) return true;
  // A redirect from service.svc is a bounce to the sign-in page.
  return response.status >= 300 && response.status < 400;
}

function retryDelay(response: Response, attempt: number): number {
  const header = response.headers.get("retry-after");
  const seconds = header ? Number(header) : NaN;
  const ms = Number.isFinite(seconds) ? seconds * 1000 : 1000 * 2 ** attempt;
  return Math.min(Math.max(ms, 250), MAX_RETRY_AFTER_MS);
}

async function parseResponse<T>(action: string, response: Response): Promise<T> {
  const text = await response.text();
  let json: any;
  try {
    json = text ? JSON.parse(text) : undefined;
  } catch {
    const looksLikeLogin = /<html|<!doctype/i.test(text);
    throw new OwaHttpError(
      looksLikeLogin
        ? `${action} returned an HTML page instead of JSON (HTTP ${response.status}); the session is probably not authenticated`
        : `${action} returned a non-JSON response (HTTP ${response.status}): ${truncate(text, 200)}`,
      response.status,
      action,
    );
  }
  if (!response.ok) {
    const fault = json?.Body ?? json ?? {};
    const name = fault.ExceptionName ?? fault.ResponseCode;
    const message = fault.FaultMessage ?? fault.MessageText ?? fault.Message ?? truncate(text, 200);
    const hint =
      name === "OwaMethodArgumentException" || name === "OwaSerializationException"
        ? " (the service rejected the request shape)"
        : /does not exist in the service method map/.test(String(message))
          ? " (this action is not available on this Outlook endpoint)"
          : "";
    throw new OwaHttpError(`${action} failed with HTTP ${response.status}${name ? ` ${name}` : ""}: ${message}${hint}`, response.status, action, name);
  }
  return json as T;
}

/** Body.ResponseMessages.Items, or an empty array. */
export function responseMessages(result: any): Json[] {
  const items = result?.Body?.ResponseMessages?.Items;
  return Array.isArray(items) ? items : [];
}

export function isErrorMessage(message: Json | undefined): boolean {
  return message?.ResponseClass === "Error";
}

export function assertSuccess(action: string, message: Json | undefined): void {
  if (!message) throw new OwaResponseError(action, "ErrorEmptyResponse", "The service returned no response message");
  if (isErrorMessage(message)) {
    throw new OwaResponseError(action, String(message.ResponseCode ?? "Error"), message.MessageText as string | undefined);
  }
}

/** Also check responses that put ResponseClass directly on Body (FindPeople, FindConversation). */
export function assertBodySuccess(action: string, result: any): any {
  const body = result?.Body;
  if (!body) throw new OwaResponseError(action, "ErrorEmptyResponse", "The service returned no body");
  if (body.ResponseClass === "Error") {
    throw new OwaResponseError(action, String(body.ResponseCode ?? "Error"), body.MessageText);
  }
  return body;
}

function fileNameFromDisposition(header: string | null): string | undefined {
  if (!header) return undefined;
  const star = /filename\*\s*=\s*(?:UTF-8'')?([^;]+)/i.exec(header);
  if (star) {
    try {
      return decodeURIComponent(star[1]!.trim().replace(/^"|"$/g, ""));
    } catch {
      /* fall through */
    }
  }
  return /filename\s*=\s*"?([^";]+)"?/i.exec(header)?.[1]?.trim();
}

async function drain(response: Response): Promise<void> {
  try {
    await response.arrayBuffer();
  } catch {
    /* ignore */
  }
}

function truncate(text: string, max: number): string {
  const clean = text.replace(/\s+/g, " ").trim();
  return clean.length > max ? `${clean.slice(0, max)}…` : clean;
}
