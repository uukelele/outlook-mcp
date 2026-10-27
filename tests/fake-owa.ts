import type { AuthProvider } from "../src/auth/providers.js";
import { loadConfig, type Config } from "../src/config.js";
import type { FetchLike } from "../src/owa/client.js";

export interface RecordedCall {
  url: string;
  action: string;
  headers: Record<string, string>;
  body: any;
}

type Handler = (body: any, call: RecordedCall) => unknown;

/** A fake OWA endpoint that routes by the Action header and records every request. */
export function fakeOwa(handlers: Record<string, Handler>) {
  const calls: RecordedCall[] = [];
  const fetch: FetchLike = async (url, init) => {
    const headers = (init.headers ?? {}) as Record<string, string>;
    const action = headers.Action ?? (url.includes("GetFileAttachment") ? "GetFileAttachment" : "");
    const body = typeof init.body === "string" ? JSON.parse(init.body) : undefined;
    const call: RecordedCall = { url, action, headers, body };
    calls.push(call);
    const handler = handlers[action];
    if (!handler) {
      return new Response(JSON.stringify({ Body: { ExceptionName: "NotHandled", FaultMessage: `no fake for ${action}` } }), { status: 500 });
    }
    const result = handler(body, call);
    if (result instanceof Response) return result;
    return new Response(JSON.stringify(result), { status: 200, headers: { "content-type": "application/json" } });
  };
  const of = (action: string) => calls.filter((c) => c.action === action);
  return { fetch, calls, of };
}

export const messages = (...items: unknown[]) => ({ Body: { ResponseMessages: { Items: items } } });
export const success = (extra: Record<string, unknown> = {}) => ({ ResponseClass: "Success", ResponseCode: "NoError", ...extra });
export const failure = (code: string, text?: string) => ({ ResponseClass: "Error", ResponseCode: code, MessageText: text });

export const testAuth: AuthProvider = {
  description: "test token",
  canRefresh: true,
  headers: async () => ({ Authorization: "Bearer test-token" }),
  identity: () => ({ upn: "me@contoso.com", puid: "PUID1", tenantId: "TENANT1" }),
};

export function testConfig(overrides: Partial<Config> = {}): Config {
  return { ...loadConfig({ OUTLOOK_ACCESS_TOKEN: "test" }), ...overrides };
}

/** Synthetic, obviously-fake Exchange item IDs (never reuse real ones: they embed the real mailbox GUID). */
export const ID = (n: number | string) =>
  Buffer.from(`SYNTHETIC-TEST-ITEM-ID-${String(n).padStart(12, "0")}`, "utf8").toString("base64");

/** A minimal FindFolder response describing a small mailbox. */
export const FOLDERS = {
  root: { Id: ID("root"), name: "Top of Information Store" },
  inbox: { Id: ID("inbox"), name: "Inbox" },
  projects: { Id: ID("proj"), name: "Projects" },
  sent: { Id: ID("sent"), name: "Sent Items" },
  archive: { Id: ID("arch"), name: "Archive" },
  calendar: { Id: ID("cal"), name: "Calendar" },
};

export function findFolderResponse() {
  const folder = (f: { Id: string; name: string }, parent: string, extra: Record<string, unknown> = {}) => ({
    __type: "Folder:#Exchange",
    FolderId: { Id: f.Id },
    ParentFolderId: { Id: parent },
    DisplayName: f.name,
    FolderClass: "IPF.Note",
    TotalCount: 10,
    UnreadCount: 2,
    ...extra,
  });
  return messages(
    success({
      RootFolder: {
        Folders: [
          folder(FOLDERS.inbox, FOLDERS.root.Id, { DistinguishedFolderId: "inbox" }),
          folder(FOLDERS.projects, FOLDERS.inbox.Id, { TotalCount: 3 }),
          folder(FOLDERS.sent, FOLDERS.root.Id, { DistinguishedFolderId: "sentitems" }),
          folder(FOLDERS.archive, FOLDERS.root.Id, { DistinguishedFolderId: "archive", TotalCount: 0 }),
          folder(FOLDERS.calendar, FOLDERS.root.Id, { DistinguishedFolderId: "calendar", FolderClass: "IPF.Appointment" }),
          folder({ Id: ID("sync"), name: "Sync Issues" }, FOLDERS.root.Id, { DistinguishedFolderId: "syncissues" }),
        ],
      },
    }),
  );
}
