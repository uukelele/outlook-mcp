export interface TokenClaims {
  aud?: string;
  exp?: number;
  upn?: string;
  unique_name?: string;
  preferred_username?: string;
  name?: string;
  puid?: string;
  tid?: string;
  scp?: string;
  appid?: string;
  [claim: string]: unknown;
}

/** Decode a JWT payload without verifying it. Returns undefined for opaque or malformed tokens. */
export function decodeJwt(token: string): TokenClaims | undefined {
  const parts = token.split(".");
  if (parts.length !== 3) return undefined;
  try {
    const json = Buffer.from(parts[1]!, "base64url").toString("utf8");
    const claims = JSON.parse(json);
    return claims && typeof claims === "object" ? (claims as TokenClaims) : undefined;
  } catch {
    return undefined;
  }
}

const GRAPH_AUDIENCES = new Set(["https://graph.microsoft.com", "00000003-0000-0000-c000-000000000000"]);
const OUTLOOK_AUDIENCES = new Set([
  "https://outlook.office.com",
  "https://outlook.office365.com",
  "https://outlook.cloud.microsoft",
  "00000002-0000-0ff1-ce00-000000000000",
]);

export type AudienceCheck = "ok" | "graph" | "unknown";

export function checkAudience(claims: TokenClaims | undefined): AudienceCheck {
  const aud = typeof claims?.aud === "string" ? claims.aud.replace(/\/+$/, "") : undefined;
  if (!aud) return "unknown";
  if (OUTLOOK_AUDIENCES.has(aud)) return "ok";
  if (GRAPH_AUDIENCES.has(aud)) return "graph";
  return "unknown";
}

export function userPrincipal(claims: TokenClaims | undefined): string | undefined {
  const value = claims?.upn ?? claims?.preferred_username ?? claims?.unique_name;
  return typeof value === "string" && value.includes("@") ? value : undefined;
}
