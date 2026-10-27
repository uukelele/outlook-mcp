import { mkdir, stat, writeFile } from "node:fs/promises";
import { dirname, isAbsolute, join } from "node:path";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import { OutlookError } from "../errors.js";
import {
  addresses,
  attachmentInput,
  attachments,
  bodyFormat,
  defineTool,
  folderRef,
  maxBodyChars,
  messageIds,
  type ToolRegistry,
} from "./helpers.js";

const importance = z.enum(["low", "normal", "high"]).describe("Message importance");

export function registerMailTools(registry: ToolRegistry): void {
  const { mail, folders, attachments: files } = registry.services;

  // ---------------------------------------------------------------- folders

  defineTool(registry, {
    name: "list_folders",
    title: "List folders",
    kind: "read",
    description:
      "List mailbox folders with their paths, IDs, well-known names and item/unread counts. " +
      "Use the returned path or ID wherever a folder is expected.",
    input: {
      kind: z.enum(["mail", "calendar", "contacts", "tasks", "all"]).optional().describe('Folder type to list (default "mail")'),
      refresh: z.boolean().optional().describe("Bypass the five-minute folder cache"),
    },
    run: async ({ kind, refresh }) => ({ folders: await folders.list(kind ?? "mail", refresh) }),
  });

  defineTool(registry, {
    name: "create_folder",
    title: "Create folder",
    kind: "write",
    description: "Create a mail folder, optionally inside another folder.",
    input: {
      name: z.string().min(1).describe("New folder name"),
      parent: folderRef("Parent folder (default: top level of the mailbox)").optional(),
    },
    run: async ({ name, parent }) => ({ created: await folders.create(name, parent ?? "msgfolderroot") }),
  });

  defineTool(registry, {
    name: "rename_folder",
    title: "Rename folder",
    kind: "write",
    description: "Rename a folder.",
    input: { folder: folderRef("Folder to rename"), name: z.string().min(1).describe("New name") },
    run: async ({ folder, name }) => ({ renamed: await folders.rename(folder, name) }),
  });

  defineTool(registry, {
    name: "move_folder",
    title: "Move folder",
    kind: "write",
    description: "Move a folder (with its contents) into another folder.",
    input: { folder: folderRef("Folder to move"), destination: folderRef("New parent folder") },
    run: async ({ folder, destination }) => {
      await folders.move(folder, destination);
      return { moved: true, folder, destination };
    },
  });

  defineTool(registry, {
    name: "delete_folder",
    title: "Delete folder",
    kind: "delete",
    description: "Delete a folder and everything in it. By default it is moved to Deleted Items; built-in folders cannot be deleted.",
    input: {
      folder: folderRef("Folder to delete"),
      permanent: z.boolean().optional().describe("Permanently delete instead of moving to Deleted Items (default false)"),
    },
    run: async ({ folder, permanent }) => {
      await folders.delete(folder, Boolean(permanent));
      return { deleted: true, folder, permanent: Boolean(permanent) };
    },
  });

  // ---------------------------------------------------------------- reading

  defineTool(registry, {
    name: "list_messages",
    title: "List messages",
    kind: "read",
    description:
      "List messages in a folder, newest first, with sender, subject, preview, read/flag state and IDs. " +
      "Page with offset/nextOffset. Times are UTC.",
    input: {
      folder: folderRef("Folder (default inbox)").optional(),
      filter: z.enum(["all", "unread", "flagged", "toOrCcMe"]).optional().describe('Server-side filter (default "all")'),
      limit: z.number().int().min(1).max(200).optional().describe("Messages to return (default 25)"),
      offset: z.number().int().min(0).optional().describe("Number of messages to skip"),
      order: z.enum(["newest", "oldest"]).optional().describe('Sort order by received time (default "newest")'),
    },
    run: (args) => mail.list(args),
  });

  defineTool(registry, {
    name: "search_messages",
    title: "Search messages",
    kind: "read",
    description:
      "Find messages by text, sender, recipient, subject, date, category, read/flag state or attachments. " +
      "Works by scanning the newest messages of each folder (scanLimit per folder) and matching subject, sender, recipient display names and the ~250-character preview; " +
      "it does not search full bodies or attachment contents. All given criteria must match. Searches all mail folders except Deleted Items, Junk, Drafts and Outbox unless folders is given.",
    input: {
      text: z.string().optional().describe('Words that must all appear (subject, preview, sender or recipients); use "quotes" for phrases'),
      from: z.string().optional().describe("Sender name or address contains this"),
      to: z.string().optional().describe("Recipient display names contain this"),
      subject: z.string().optional().describe("Subject contains this"),
      folders: z.array(z.string()).optional().describe('Folders to search (names, paths or IDs), or ["all"]'),
      after: z.string().optional().describe("Only messages received at/after this date or time (ISO 8601)"),
      before: z.string().optional().describe("Only messages received before this date or time (ISO 8601)"),
      unread: z.boolean().optional().describe("Only unread messages"),
      flagged: z.boolean().optional().describe("Only flagged messages"),
      hasAttachments: z.boolean().optional().describe("Only messages with (true) or without (false) attachments"),
      category: z.string().optional().describe("Only messages with this category"),
      limit: z.number().int().min(1).max(200).optional().describe("Maximum results (default 25)"),
      scanLimit: z.number().int().min(50).max(10_000).optional().describe("Newest messages to scan per folder (default 1000)"),
    },
    run: (args) => mail.search(args),
  });

  defineTool(registry, {
    name: "get_message",
    title: "Read message",
    kind: "read",
    description: "Get a message with full body (a string; see bodyFormat), recipients and file attachments (inline images are only counted; list_attachments shows them). Does not mark it read unless markAsRead is true.",
    input: {
      id: z.string().min(1).describe("Message ID"),
      format: bodyFormat.optional(),
      maxBodyChars: maxBodyChars.optional(),
      markAsRead: z.boolean().optional().describe("Also mark the message as read (default false)"),
    },
    run: ({ id, format, maxBodyChars, markAsRead }) => mail.get(id, { format, maxBodyChars: maxBodyChars ?? 20_000, markAsRead }),
  });

  defineTool(registry, {
    name: "get_thread",
    title: "Read conversation",
    kind: "read",
    description:
      "Get every message in a conversation (across folders), oldest first. Each body contains only the new text of that message, without quoted history.",
    input: {
      messageId: z.string().optional().describe("Any message in the conversation"),
      conversationId: z.string().optional().describe("Conversation ID (alternative to messageId)"),
      maxBodyChars: z.number().int().min(200).max(100_000).optional().describe("Per-message body limit (default 4000)"),
    },
    run: ({ messageId, conversationId, maxBodyChars }) => mail.thread({ messageId, conversationId }, { maxBodyChars }),
  });

  // ---------------------------------------------------------------- composing

  const compose = {
    to: addresses("Recipients").optional(),
    cc: addresses("Cc recipients").optional(),
    bcc: addresses("Bcc recipients").optional(),
    subject: z.string().optional().describe("Subject line"),
    body: z.string().optional().describe("Message body"),
    bodyFormat: bodyFormat.optional(),
    importance: importance.optional(),
  };

  defineTool(registry, {
    name: "send_message",
    title: "Send email",
    kind: "send",
    description: "Compose and immediately send a new email (a copy is saved in Sent Items). Use create_draft instead if the user wants to review it first.",
    input: { ...compose, to: addresses("Recipients"), subject: z.string().describe("Subject line"), body: z.string().describe("Message body"), attachments: attachments.optional() },
    run: (args) => mail.send(args),
  });

  defineTool(registry, {
    name: "create_draft",
    title: "Create draft",
    kind: "write",
    description: "Save a new email in Drafts without sending it. Returns the draft ID for update_draft, add_attachment and send_draft.",
    input: { ...compose, attachments: attachments.optional() },
    run: (args) => mail.createDraft(args),
  });

  defineTool(registry, {
    name: "update_draft",
    title: "Update draft",
    kind: "write",
    description: "Change a draft's subject, body, importance or recipients. Recipient lists replace the existing ones; an empty list clears them.",
    input: { id: z.string().min(1).describe("Draft ID"), ...compose },
    run: ({ id, ...rest }) => mail.updateDraft(id, rest),
  });

  defineTool(registry, {
    name: "send_draft",
    title: "Send draft",
    kind: "send",
    description: "Send an existing draft (including reply/forward drafts).",
    input: { id: z.string().min(1).describe("Draft ID") },
    run: ({ id }) => mail.sendDraft(id),
  });

  defineTool(registry, {
    name: "reply_message",
    title: "Reply to email",
    kind: "send",
    description:
      "Reply (or reply-all) to a message. The original is quoted below your text. Sends immediately unless send is false, in which case a draft is saved and its ID returned.",
    input: {
      id: z.string().min(1).describe("ID of the message to reply to"),
      body: z.string().describe("Reply text (only the new part; the original is quoted automatically)"),
      bodyFormat: bodyFormat.optional(),
      replyAll: z.boolean().optional().describe("Reply to all recipients (default false)"),
      send: z.boolean().optional().describe("Send now (default true) or save as draft (false)"),
      attachments: attachments.optional(),
    },
    run: ({ id, send, ...rest }) => mail.reply(id, { ...rest, send: send !== false }),
  });

  defineTool(registry, {
    name: "forward_message",
    title: "Forward email",
    kind: "send",
    description: "Forward a message (with its attachments) to new recipients. Sends immediately unless send is false.",
    input: {
      id: z.string().min(1).describe("ID of the message to forward"),
      to: addresses("Recipients"),
      cc: addresses("Cc recipients").optional(),
      bcc: addresses("Bcc recipients").optional(),
      body: z.string().optional().describe("Text to add above the forwarded message"),
      bodyFormat: bodyFormat.optional(),
      send: z.boolean().optional().describe("Send now (default true) or save as draft (false)"),
      attachments: attachments.optional(),
    },
    run: ({ id, send, ...rest }) => mail.forward(id, { ...rest, send: send !== false }),
  });

  // ---------------------------------------------------------------- triage

  defineTool(registry, {
    name: "mark_messages_read",
    title: "Mark read/unread",
    kind: "write",
    description: "Mark messages as read or unread (no read receipts are sent).",
    input: { ids: messageIds, read: z.boolean().describe("true = read, false = unread") },
    run: ({ ids, read }) => mail.markRead(ids, read),
  });

  defineTool(registry, {
    name: "flag_messages",
    title: "Flag messages",
    kind: "write",
    description: "Set the follow-up flag on messages.",
    input: { ids: messageIds, status: z.enum(["flagged", "complete", "none"]).describe('"flagged", "complete" or "none" (clear)') },
    run: ({ ids, status }) => mail.setFlag(ids, status),
  });

  defineTool(registry, {
    name: "categorize_messages",
    title: "Categorize messages",
    kind: "write",
    description: "Set, add or remove categories on messages. See list_categories for the mailbox's category names.",
    input: {
      ids: messageIds,
      categories: z.array(z.string()).describe("Category names"),
      mode: z.enum(["set", "add", "remove"]).optional().describe('"add" (default) keeps existing categories, "set" replaces them, "remove" removes the given ones'),
    },
    run: ({ ids, categories, mode }) => mail.setCategories(ids, categories, mode ?? "add"),
  });

  defineTool(registry, {
    name: "move_messages",
    title: "Move messages",
    kind: "write",
    description: "Move messages to another folder (e.g. archive). Message IDs stay valid after moving.",
    input: { ids: messageIds, destination: folderRef("Destination folder") },
    run: ({ ids, destination }) => mail.move(ids, destination),
  });

  defineTool(registry, {
    name: "copy_messages",
    title: "Copy messages",
    kind: "write",
    description: "Copy messages to another folder; returns the IDs of the copies.",
    input: { ids: messageIds, destination: folderRef("Destination folder") },
    run: ({ ids, destination }) => mail.copy(ids, destination),
  });

  defineTool(registry, {
    name: "delete_messages",
    title: "Delete messages",
    kind: "delete",
    description: "Delete messages. By default they are moved to Deleted Items (recoverable); permanent deletion cannot be undone.",
    input: { ids: messageIds, permanent: z.boolean().optional().describe("Permanently delete (default false)") },
    run: ({ ids, permanent }) => mail.delete(ids, Boolean(permanent)),
  });

  // ---------------------------------------------------------------- attachments

  defineTool(registry, {
    name: "list_attachments",
    title: "List attachments",
    kind: "read",
    description: "List the attachments of a message, event or contact.",
    input: { id: z.string().min(1).describe("Message, event or contact ID") },
    run: async ({ id }) => ({ attachments: await files.list(id) }),
  });

  defineTool(registry, {
    name: "download_attachment",
    title: "Download attachment",
    kind: "read",
    description:
      "Download a file attachment. With savePath the file is written to disk; otherwise text files are returned as text, images as images, and other files as an embedded base64 resource (up to 10 MB).",
    input: {
      id: z.string().min(1).describe("ID of the message/event that has the attachment"),
      attachmentId: z.string().optional().describe("Attachment ID from get_message/list_attachments"),
      name: z.string().optional().describe("Attachment file name (alternative to attachmentId)"),
      savePath: z.string().optional().describe("Absolute path to save the file to (a directory or full file path)"),
      overwrite: z.boolean().optional().describe("Allow replacing an existing file at savePath (default false)"),
    },
    run: async ({ id, attachmentId, name, savePath, overwrite }) => {
      const file = await files.download(id, { attachmentId, name });
      const meta = { name: file.name, contentType: file.contentType, size: file.data.length };
      if (savePath) {
        const path = await chooseSavePath(savePath, file.name ?? "attachment", Boolean(overwrite));
        await mkdir(dirname(path), { recursive: true });
        await writeFile(path, file.data, { flag: overwrite ? "w" : "wx" });
        return { ...meta, savedTo: path };
      }
      return inlineAttachment(file.name ?? "attachment", file.contentType ?? "application/octet-stream", file.data);
    },
  });

  defineTool(registry, {
    name: "add_attachment",
    title: "Add attachment",
    kind: "write",
    description: "Attach a file to a draft (or to an event/contact you own).",
    input: { id: z.string().min(1).describe("Draft, event or contact ID"), attachment: attachmentInput },
    run: async ({ id, attachment }) => ({ added: await files.add(id, attachment) }),
  });

  defineTool(registry, {
    name: "remove_attachment",
    title: "Remove attachment",
    kind: "delete",
    description: "Remove an attachment from a draft (or from an event/contact you own).",
    input: { attachmentId: z.string().min(1).describe("Attachment ID") },
    run: async ({ attachmentId }) => {
      await files.remove(attachmentId);
      return { removed: true };
    },
  });
}

const INLINE_TEXT_LIMIT = 1024 * 1024;
const INLINE_BINARY_LIMIT = 10 * 1024 * 1024;

function isTextual(contentType: string): boolean {
  return /^text\/|\/(json|xml|csv|javascript|x-yaml|yaml|calendar)\b|\+xml|\+json/i.test(contentType);
}

function inlineAttachment(name: string, contentType: string, data: Buffer): CallToolResult {
  const header = { type: "text" as const, text: JSON.stringify({ name, contentType, size: data.length }) };
  const mime = contentType.split(";")[0]!.trim().toLowerCase();
  if (isTextual(mime) && data.length <= INLINE_TEXT_LIMIT) {
    return { content: [header, { type: "text", text: data.toString("utf8") }] };
  }
  if (data.length > INLINE_BINARY_LIMIT) {
    throw new OutlookError(`"${name}" is ${data.length} bytes; pass savePath to save it to disk instead`);
  }
  if (/^image\/(png|jpeg|gif|webp)$/.test(mime)) {
    return { content: [header, { type: "image", data: data.toString("base64"), mimeType: mime }] };
  }
  return {
    content: [
      header,
      { type: "resource", resource: { uri: `attachment:///${encodeURIComponent(name)}`, mimeType: mime, blob: data.toString("base64") } },
    ],
  };
}

async function chooseSavePath(savePath: string, fileName: string, overwrite: boolean): Promise<string> {
  if (!isAbsolute(savePath)) throw new OutlookError("savePath must be an absolute path");
  const safeName = fileName.replace(/[/\\:]/g, "_").replace(/\0/g, "");
  let target = savePath;
  const directory = /[\\/]$/.test(savePath) || (await stat(savePath).then((s) => s.isDirectory(), () => false));
  if (directory) target = join(savePath, safeName);
  if (!overwrite) {
    const exists = await stat(target).then(
      () => true,
      () => false,
    );
    if (exists) throw new OutlookError(`${target} already exists; pass overwrite: true to replace it`);
  }
  return target;
}
