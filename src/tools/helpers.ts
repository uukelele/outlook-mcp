import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult, ToolAnnotations } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { errorMessage } from "../errors.js";
import type { AttachmentService } from "../services/attachments.js";
import type { CalendarService } from "../services/calendar.js";
import type { ContactService } from "../services/contacts.js";
import type { FolderService } from "../services/folders.js";
import type { MailService } from "../services/mail.js";
import type { MailboxService } from "../services/mailbox.js";

export interface Services {
  mailbox: MailboxService;
  folders: FolderService;
  mail: MailService;
  attachments: AttachmentService;
  calendar: CalendarService;
  contacts: ContactService;
}

export interface ToolRegistry {
  server: McpServer;
  services: Services;
  readOnly: boolean;
  registered: string[];
}

type Shape = z.ZodRawShape;

interface ToolSpec<S extends Shape> {
  name: string;
  title: string;
  description: string;
  input: S;
  /** read: no side effects; write: modifies the mailbox; send: sends mail or invitations to other people. */
  kind: "read" | "write" | "delete" | "send";
  run: (args: z.infer<z.ZodObject<S>>) => Promise<unknown | CallToolResult>;
}

const ANNOTATIONS: Record<ToolSpec<Shape>["kind"], ToolAnnotations> = {
  read: { readOnlyHint: true, openWorldHint: false },
  write: { readOnlyHint: false, destructiveHint: false, openWorldHint: false },
  delete: { readOnlyHint: false, destructiveHint: true, openWorldHint: false },
  send: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
};

export function defineTool<S extends Shape>(registry: ToolRegistry, spec: ToolSpec<S>): void {
  if (registry.readOnly && spec.kind !== "read") return;
  registry.server.registerTool(
    spec.name,
    {
      title: spec.title,
      description: spec.description,
      inputSchema: spec.input,
      annotations: { title: spec.title, ...ANNOTATIONS[spec.kind] },
    },
    (async (args: z.infer<z.ZodObject<S>>) => {
      try {
        const result = await spec.run(args);
        return isToolResult(result) ? result : json(result);
      } catch (error) {
        return { isError: true, content: [{ type: "text", text: `Error: ${errorMessage(error)}` }] } satisfies CallToolResult;
      }
    }) as any,
  );
  registry.registered.push(spec.name);
}

export function json(data: unknown): CallToolResult {
  return { content: [{ type: "text", text: JSON.stringify(data, null, 2) }] };
}

function isToolResult(value: unknown): value is CallToolResult {
  return Boolean(value && typeof value === "object" && Array.isArray((value as CallToolResult).content));
}

// ---------------------------------------------------------------------------
// Shared schemas

export const addresses = (what: string) =>
  z
    .array(z.string().min(3))
    .describe(`${what}: email addresses, either "jane@contoso.com" or "Jane Doe <jane@contoso.com>"`);

export const messageIds = z.array(z.string().min(1)).min(1).max(500).describe("Message IDs (from list_messages, search_messages or get_message)");

export const bodyFormat = z.enum(["text", "html"]).describe('Body format: "text" (default) or "html"');

export const folderRef = (description: string) =>
  z
    .string()
    .min(1)
    .describe(
      `${description}. Accepts a well-known name (inbox, drafts, sentitems/sent, deleteditems/trash, junkemail/junk, archive, outbox), a display name or path such as "Inbox/Projects", or a folder ID from list_folders`,
    );

export const attachmentInput = z
  .object({
    path: z.string().optional().describe("Absolute path of a local file to attach"),
    contentBase64: z.string().optional().describe("File content as base64 (alternative to path)"),
    text: z.string().optional().describe("Plain-text content (alternative to path)"),
    name: z.string().optional().describe("File name shown to recipients; required unless path is given"),
    contentType: z.string().optional().describe("MIME type; guessed from the file name if omitted"),
  })
  .describe("Exactly one of path, contentBase64 or text");

export const attachments = z.array(attachmentInput).max(20).describe("Files to attach");

export const maxBodyChars = z
  .number()
  .int()
  .min(200)
  .max(500_000)
  .describe("Truncate the body after this many characters (default 20000)");
