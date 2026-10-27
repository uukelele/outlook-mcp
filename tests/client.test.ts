import { describe, expect, it, vi } from "vitest";
import type { AuthProvider } from "../src/auth/providers.js";
import { AuthError, OwaHttpError, OwaResponseError } from "../src/errors.js";
import { OwaClient } from "../src/owa/client.js";
import { failure, fakeOwa, messages, success, testAuth } from "./fake-owa.js";

function client(fetch: any, auth: AuthProvider = testAuth) {
  return new OwaClient({
    origin: "https://outlook.cloud.microsoft",
    attachmentsOrigin: "https://attachments.office.net",
    auth,
    userAgent: "test",
    requestTimeoutMs: 5000,
    fetch,
    sleep: async () => {},
  });
}

describe("OwaClient", () => {
  it("sends OWA headers, the action and an enveloped JSON body", async () => {
    const owa = fakeOwa({ FindItem: () => messages(success()) });
    await client(owa.fetch).call("FindItem", { Traversal: "Shallow" });
    const [call] = owa.calls;
    expect(call!.url).toMatch(/^https:\/\/outlook\.cloud\.microsoft\/owa\/service\.svc\?action=FindItem&app=Mail&n=1$/);
    expect(call!.headers).toMatchObject({
      Authorization: "Bearer test-token",
      Action: "FindItem",
      "Content-Type": "application/json; charset=utf-8",
      "X-Requested-With": "XMLHttpRequest",
      Origin: "https://outlook.cloud.microsoft",
      "X-AnchorMailbox": "PUID:PUID1@TENANT1",
    });
    expect(call!.headers.Prefer).toContain('IdType="ImmutableId"');
    expect(call!.body.__type).toBe("FindItemJsonRequest:#Exchange");
    expect(call!.body.Body).toEqual({ __type: "FindItemRequest:#Exchange", Traversal: "Shallow" });
  });

  it("refreshes credentials once after a 401", async () => {
    let calls = 0;
    const headers = vi.fn(async (opts?: { forceRefresh?: boolean }) => ({ Authorization: opts?.forceRefresh ? "Bearer new" : "Bearer old" }));
    const auth = { ...testAuth, headers };
    const fetch = async (_url: string, init: RequestInit) => {
      calls++;
      const auth = (init.headers as Record<string, string>).Authorization;
      return auth === "Bearer new" ? new Response(JSON.stringify(messages(success())), { status: 200 }) : new Response("", { status: 401 });
    };
    await client(fetch, auth).call("GetItem", {});
    expect(calls).toBe(2);
    expect(headers).toHaveBeenLastCalledWith({ forceRefresh: true });
  });

  it("gives up with AuthError when refreshed credentials are also rejected", async () => {
    const fetch = async () => new Response("", { status: 401 });
    await expect(client(fetch).call("GetItem", {})).rejects.toBeInstanceOf(AuthError);
  });

  it("treats redirects as authentication failures (sign-in bounce)", async () => {
    const fetch = async () => new Response("", { status: 302, headers: { location: "https://login.microsoftonline.com/" } });
    await expect(client(fetch).call("GetItem", {})).rejects.toBeInstanceOf(AuthError);
  });

  it("retries throttled requests", async () => {
    let calls = 0;
    const fetch = async () =>
      ++calls < 3 ? new Response("", { status: 429, headers: { "retry-after": "1" } }) : new Response(JSON.stringify({ ok: 1 }), { status: 200 });
    expect(await client(fetch).call("GetItem", {})).toEqual({ ok: 1 });
    expect(calls).toBe(3);
  });

  it("reports HTML login pages and service faults clearly", async () => {
    const html = async () => new Response("<!DOCTYPE html><html>sign in</html>", { status: 200 });
    await expect(client(html).call("GetItem", {})).rejects.toThrow(/HTML page instead of JSON/);

    const fault = async () =>
      new Response(JSON.stringify({ Body: { ExceptionName: "OwaMethodArgumentException", FaultMessage: "Invalid argument" } }), { status: 400 });
    const error = await client(fault).call("FindItem", {}).catch((e) => e);
    expect(error).toBeInstanceOf(OwaHttpError);
    expect(error.message).toMatch(/HTTP 400 OwaMethodArgumentException: Invalid argument \(the service rejected the request shape\)/);
  });

  it("raises OwaResponseError for HTTP 200 responses with ResponseClass Error", async () => {
    const owa = fakeOwa({ GetItem: () => messages(failure("ErrorItemNotFound", "The specified object was not found in the store.")) });
    const error = await client(owa.fetch).callForMessages("GetItem", {}).catch((e) => e);
    expect(error).toBeInstanceOf(OwaResponseError);
    expect(error.responseCode).toBe("ErrorItemNotFound");
    expect(error.message).toMatch(/moved or deleted/);
  });

  it("sends raw payloads untouched", async () => {
    const owa = fakeOwa({ GetAttachmentDownloadToken: () => "jwt-token" });
    const payload = { __type: "GetAttachmentDownloadTokenRequest:#Exchange", UseNewDownloadTokenLifetime: true };
    expect(await client(owa.fetch).call("GetAttachmentDownloadToken", payload, { raw: true })).toBe("jwt-token");
    expect(owa.calls[0]!.body).toEqual(payload);
  });
});
