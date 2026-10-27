/**
 * Live, read-only tests against the real mailbox configured in the environment.
 * Run with: npm run test:live
 *
 * The server is created in read-only mode, so write tools are not even registered;
 * nothing in this file can send, create, change or delete anything.
 */
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { beforeAll, describe, expect, it } from "vitest";
import { loadConfig } from "../src/config.js";
import { createServer } from "../src/server.js";

const LIVE = process.env.OUTLOOK_LIVE_TESTS === "1";

describe.runIf(LIVE)("live (read-only)", { timeout: 120_000 }, () => {
  let client: Client;
  let tools: string[];

  async function call(name: string, args: Record<string, unknown> = {}) {
    const result = await client.callTool({ name, arguments: args });
    const content = result.content as { type: string; text?: string }[];
    if (result.isError) throw new Error(`${name} failed: ${content[0]?.text}`);
    return { result, content, data: content[0]?.type === "text" ? safeJson(content[0].text!) : undefined };
  }

  beforeAll(async () => {
    const { server, tools: registered } = createServer({ ...loadConfig(), readOnly: true });
    tools = registered;
    const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
    await server.connect(serverSide);
    client = new Client({ name: "live-test", version: "1.0.0" });
    await client.connect(clientSide);
  });

  it("only exposes read tools", () => {
    expect(tools).not.toContain("send_message");
    expect(tools).not.toContain("delete_messages");
    expect(tools).toHaveLength(15);
  });

  it("get_mailbox_info / list_categories", async () => {
    const { data: info } = await call("get_mailbox_info");
    expect(info.email).toMatch(/@/);
    expect(info.timeZone).toBeTruthy();
    expect(info.folders.inbox.total).toBeGreaterThanOrEqual(0);
    const { data } = await call("list_categories");
    expect(Array.isArray(data.categories)).toBe(true);
  });

  it("list_folders", async () => {
    const { data } = await call("list_folders");
    const inbox = data.folders.find((f: any) => f.wellKnownName === "inbox");
    expect(inbox?.path).toBe(inbox?.name);
    expect(data.folders.every((f: any) => typeof f.path === "string" && f.id)).toBe(true);
  });

  it("list_messages with paging, filters and order", async () => {
    const { data: page1 } = await call("list_messages", { limit: 3 });
    expect(page1.messages.length).toBeGreaterThan(0);
    const dates = page1.messages.map((m: any) => m.receivedAt);
    expect([...dates].sort().reverse()).toEqual(dates);
    if (page1.nextOffset !== undefined) {
      const { data: page2 } = await call("list_messages", { limit: 3, offset: page1.nextOffset });
      expect(page2.messages[0]?.id).not.toBe(page1.messages[0].id);
    }
    const { data: unread } = await call("list_messages", { filter: "unread", limit: 5 });
    expect(unread.messages.every((m: any) => m.isRead === false)).toBe(true);
    const { data: oldest } = await call("list_messages", { order: "oldest", limit: 1 });
    expect(oldest.messages[0].receivedAt <= page1.messages[0].receivedAt).toBe(true);
    await call("list_messages", { folder: "sent", limit: 2 });
  });

  it("search_messages finds a known message", async () => {
    const { data: list } = await call("list_messages", { limit: 10 });
    const target = list.messages.find((m: any) => /[A-Za-z]{5,}/.test(m.subject));
    expect(target).toBeTruthy();
    const word = /[A-Za-z]{5,}/.exec(target.subject)![0];
    const { data } = await call("search_messages", { text: word, folders: ["inbox"], scanLimit: 200 });
    expect(data.messages.some((m: any) => m.id === target.id)).toBe(true);
    expect(data.messages.every((m: any) => JSON.stringify(m).toLowerCase().includes(word.toLowerCase()))).toBe(true);
  });

  it("get_message (text and html) and get_thread", async () => {
    const { data: list } = await call("list_messages", { limit: 1 });
    const id = list.messages[0].id;
    const { data: message } = await call("get_message", { id, maxBodyChars: 2000 });
    expect(message.id).toBe(id);
    expect(message.bodyFormat).toBe("text");
    expect(typeof message.body).toBe("string");
    const { data: html } = await call("get_message", { id, format: "html", maxBodyChars: 2000 });
    expect(html.bodyFormat).toBe("html");
    const { data: thread } = await call("get_thread", { messageId: id });
    expect(thread.messages.some((m: any) => m.id === id)).toBe(true);
  });

  it("list_attachments / download_attachment (inline, no files written)", async () => {
    const { data: found } = await call("search_messages", { hasAttachments: true, folders: ["inbox"], limit: 10, scanLimit: 300 });
    let downloaded = false;
    for (const message of found.messages) {
      const { data } = await call("list_attachments", { id: message.id });
      const file = data.attachments
        .filter((a: any) => !a.kind && typeof a.size === "number" && a.size < 2_000_000)
        .sort((a: any, b: any) => a.size - b.size)[0];
      if (!file) continue;
      const { content } = await call("download_attachment", { id: message.id, attachmentId: file.id });
      const meta = JSON.parse(content[0]!.text!);
      expect(meta.name).toBe(file.name);
      expect(meta.size).toBeGreaterThan(0);
      expect(content).toHaveLength(2);
      downloaded = true;
      break;
    }
    expect(downloaded).toBe(true);
  });

  it("list_calendars / list_events / get_event", async () => {
    const { data: calendars } = await call("list_calendars");
    expect(calendars.calendars.some((c: any) => c.wellKnownName === "calendar")).toBe(true);
    const start = new Date(Date.now() - 14 * 86_400_000).toISOString();
    const end = new Date(Date.now() + 14 * 86_400_000).toISOString();
    const { data } = await call("list_events", { start, end, limit: 20 });
    expect(data.range.start).toMatch(/Z$/);
    const starts = data.events.map((e: any) => e.start);
    expect([...starts].sort()).toEqual(starts);
    if (data.events.length) {
      const { data: event } = await call("get_event", { id: data.events[0].id, maxBodyChars: 500 });
      expect(event.id).toBe(data.events[0].id);
      expect(event.start).toMatch(/Z$/);
    }
  });

  it("list_contacts / get_contact / search_people", async () => {
    const { data } = await call("list_contacts", { limit: 5 });
    expect(Array.isArray(data.contacts)).toBe(true);
    if (data.contacts.length) {
      const { data: contact } = await call("get_contact", { id: data.contacts[0].id });
      expect(contact.id).toBe(data.contacts[0].id);
    }
    const { data: info } = await call("get_mailbox_info");
    const query = String(info.email).split("@")[0];
    const { data: people } = await call("search_people", { query, source: "directory", limit: 5 });
    expect(people.people.some((p: any) => String(p.email).toLowerCase() === String(info.email).toLowerCase())).toBe(true);
  });
});

function safeJson(text: string): any {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}
