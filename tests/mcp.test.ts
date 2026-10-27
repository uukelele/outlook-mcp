import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { describe, expect, it } from "vitest";
import { createServer } from "../src/server.js";
import { fakeOwa, findFolderResponse, messages, success, testAuth, testConfig } from "./fake-owa.js";

async function connect(readOnly: boolean, handlers: Record<string, (body: any) => unknown> = {}) {
  const owa = fakeOwa({ FindFolder: () => findFolderResponse(), ...handlers });
  const { server } = createServer(testConfig({ readOnly }), { auth: testAuth, fetch: owa.fetch });
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  await server.connect(serverSide);
  const client = new Client({ name: "test", version: "1.0.0" });
  await client.connect(clientSide);
  return { client, owa };
}

const READ_TOOLS = [
  "get_mailbox_info",
  "list_categories",
  "list_folders",
  "list_messages",
  "search_messages",
  "get_message",
  "get_thread",
  "list_attachments",
  "download_attachment",
  "list_calendars",
  "list_events",
  "get_event",
  "list_contacts",
  "get_contact",
  "search_people",
];

describe("MCP server", () => {
  it("exposes the full tool set with annotations", async () => {
    const { client } = await connect(false);
    const { tools } = await client.listTools();
    const names = tools.map((t) => t.name);
    expect(names).toHaveLength(40);
    expect(names).toEqual(expect.arrayContaining([...READ_TOOLS, "send_message", "reply_message", "create_event", "delete_messages", "create_contact"]));
    const byName = Object.fromEntries(tools.map((t) => [t.name, t]));
    expect(byName.list_messages!.annotations).toMatchObject({ readOnlyHint: true });
    expect(byName.send_message!.annotations).toMatchObject({ readOnlyHint: false, openWorldHint: true });
    expect(byName.delete_messages!.annotations).toMatchObject({ destructiveHint: true });
    expect(byName.send_message!.inputSchema.required).toEqual(expect.arrayContaining(["to", "subject", "body"]));
  });

  it("registers only read tools in read-only mode", async () => {
    const { client } = await connect(true);
    const names = (await client.listTools()).tools.map((t) => t.name).sort();
    expect(names).toEqual([...READ_TOOLS].sort());
  });

  it("returns JSON text results", async () => {
    const { client } = await connect(true, {
      FindItem: () =>
        messages(success({ RootFolder: { Items: [{ ItemId: { Id: "M1" }, Subject: "Hi", IsRead: false }], TotalItemsInView: 1, IncludesLastItemInRange: true } })),
    });
    const result = await client.callTool({ name: "list_messages", arguments: { limit: 5 } });
    expect(result.isError).toBeFalsy();
    const data = JSON.parse((result.content as any)[0].text);
    expect(data).toMatchObject({ folder: "inbox", total: 1, messages: [{ id: "M1", subject: "Hi", isRead: false }] });
  });

  it("reports service errors as tool errors instead of failing the protocol", async () => {
    const { client } = await connect(true);
    const result = await client.callTool({ name: "list_messages", arguments: { folder: "No Such Folder" } });
    expect(result.isError).toBe(true);
    expect((result.content as any)[0].text).toMatch(/Folder "No Such Folder" was not found/);
  });

  it("validates arguments against the schema", async () => {
    const { client } = await connect(true);
    const result = await client.callTool({ name: "list_messages", arguments: { limit: 0 } });
    expect(result.isError).toBe(true);
  });
});
