import { describe, expect, it, vi } from "vitest";
import { checkAudience, decodeJwt } from "../src/auth/jwt.js";
import { AzureCliTokenSource, BearerAuthProvider, CookieAuthProvider, parseTokenOutput, type TokenSource } from "../src/auth/providers.js";
import { loadConfig } from "../src/config.js";

function jwt(claims: Record<string, unknown>): string {
  const enc = (o: object) => Buffer.from(JSON.stringify(o)).toString("base64url");
  return `${enc({ alg: "none" })}.${enc(claims)}.sig`;
}

const NOW = 1_800_000_000_000;

function source(tokens: string[], refreshable = true): TokenSource & { fetch: ReturnType<typeof vi.fn> } {
  let i = 0;
  return { description: "test source", refreshable, fetch: vi.fn(async () => ({ token: tokens[Math.min(i++, tokens.length - 1)]! })) };
}

describe("jwt", () => {
  it("decodes claims and classifies audiences", () => {
    const claims = decodeJwt(jwt({ aud: "https://outlook.office.com/", upn: "me@contoso.com" }));
    expect(claims?.upn).toBe("me@contoso.com");
    expect(checkAudience(claims)).toBe("ok");
    expect(checkAudience({ aud: "https://graph.microsoft.com" })).toBe("graph");
    expect(decodeJwt("opaque-token")).toBeUndefined();
  });
});

describe("BearerAuthProvider", () => {
  it("caches tokens until shortly before expiry", async () => {
    const exp = NOW / 1000 + 3600;
    const src = source([jwt({ aud: "https://outlook.office.com", exp, puid: "P", tid: "T", upn: "me@contoso.com" })]);
    let now = NOW;
    const provider = new BearerAuthProvider(src, () => {}, () => now);
    await provider.headers();
    await provider.headers();
    expect(src.fetch).toHaveBeenCalledTimes(1);
    expect(provider.identity()).toEqual({ upn: "me@contoso.com", name: undefined, puid: "P", tenantId: "T" });
    now = NOW + 3600_000 - 60_000; // within the 5-minute refresh window
    await provider.headers();
    expect(src.fetch).toHaveBeenCalledTimes(2);
  });

  it("refreshes on demand and de-duplicates concurrent loads", async () => {
    const src = source([jwt({ aud: "https://outlook.office.com", exp: NOW / 1000 + 3600 })]);
    const provider = new BearerAuthProvider(src, () => {}, () => NOW);
    await Promise.all([provider.headers(), provider.headers(), provider.headers()]);
    expect(src.fetch).toHaveBeenCalledTimes(1);
    await provider.headers({ forceRefresh: true });
    expect(src.fetch).toHaveBeenCalledTimes(2);
  });

  it("rejects Microsoft Graph tokens with a helpful message", async () => {
    const provider = new BearerAuthProvider(source([jwt({ aud: "https://graph.microsoft.com" })]));
    await expect(provider.headers()).rejects.toThrow(/Microsoft Graph/);
  });

  it("cannot refresh a static token after it is rejected", async () => {
    const provider = new BearerAuthProvider(source([jwt({ aud: "https://outlook.office.com" })], false));
    await provider.headers();
    await expect(provider.headers({ forceRefresh: true })).rejects.toThrow(/cannot be refreshed/);
  });
});

describe("parseTokenOutput", () => {
  it("accepts raw tokens, Bearer prefixes and JSON", () => {
    expect(parseTokenOutput("abc\n", "x")).toEqual({ token: "abc" });
    expect(parseTokenOutput("Bearer abc", "x")).toEqual({ token: "abc" });
    expect(parseTokenOutput('{"access_token":"abc","expires_on":1800000000}', "x")).toEqual({ token: "abc", expiresAt: 1_800_000_000_000 });
    expect(() => parseTokenOutput("", "x")).toThrow(/no token/);
    expect(() => parseTokenOutput('{"nope":1}', "x")).toThrow(/no accessToken/);
  });
});

describe("AzureCliTokenSource", () => {
  it("rejects shell-unsafe settings before invoking the CLI (Windows runs it through the shell)", async () => {
    const badTenant = new AzureCliTokenSource("https://outlook.office.com", "tenant & calc");
    await expect(badTenant.fetch()).rejects.toThrow(/OUTLOOK_AZ_TENANT/);
    const badResource = new AzureCliTokenSource("https://evil && whoami");
    await expect(badResource.fetch()).rejects.toThrow(/OUTLOOK_AZ_RESOURCE/);
  });
});

describe("CookieAuthProvider", () => {
  it("sends the cookie and extracts the canary", async () => {
    const provider = new CookieAuthProvider("a=1; X-OWA-CANARY=can123; b=2");
    expect(await provider.headers()).toEqual({ Cookie: "a=1; X-OWA-CANARY=can123; b=2", "X-OWA-CANARY": "can123" });
    await expect(provider.headers({ forceRefresh: true })).rejects.toThrow(/expired/);
  });
});

describe("loadConfig", () => {
  it("defaults to Azure CLI", () => {
    expect(loadConfig({}).auth).toEqual({ kind: "azure-cli", resource: "https://outlook.office.com", tenant: undefined });
  });

  it("picks the first configured credential in auto mode", () => {
    expect(loadConfig({ OUTLOOK_ACCESS_TOKEN: "t", OUTLOOK_TOKEN_COMMAND: "cmd" }).auth.kind).toBe("token");
    expect(loadConfig({ OUTLOOK_TOKEN_COMMAND: "cmd" }).auth.kind).toBe("command");
    expect(loadConfig({ OUTLOOK_AUTH: "azure-cli", OUTLOOK_ACCESS_TOKEN: "t" }).auth.kind).toBe("azure-cli");
  });

  it("validates settings", () => {
    expect(() => loadConfig({ OUTLOOK_AUTH: "token" })).toThrow(/OUTLOOK_ACCESS_TOKEN/);
    expect(() => loadConfig({ OUTLOOK_AUTH: "magic" })).toThrow(/OUTLOOK_AUTH must be one of/);
    expect(() => loadConfig({ OUTLOOK_MAX_DOWNLOAD_MB: "-1" })).toThrow(/positive number/);
    expect(loadConfig({ OUTLOOK_READ_ONLY: "true", OUTLOOK_ORIGIN: "https://example.com/" })).toMatchObject({
      readOnly: true,
      origin: "https://example.com",
    });
  });
});
