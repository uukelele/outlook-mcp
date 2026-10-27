import { OutlookError, errorMessage, describeResponseCode } from "../errors.js";
import {
  bodyContent,
  chunk,
  deleteItemField,
  folderIdRef,
  indexedPage,
  itemChange,
  itemId,
  itemShape,
  mailRecipients,
  propertyUri,
  setItemField,
  sortBy,
  t,
  targetFolder,
  type BodyFormat,
  type Json,
} from "../owa/builders.js";
import { isErrorMessage, responseMessages, assertSuccess, type OwaClient } from "../owa/client.js";
import { messageDetail, messageSummary } from "../owa/normalize.js";
import { parseRangeInput } from "../owa/time.js";
import type { AttachmentInput, AttachmentService } from "./attachments.js";
import type { FolderService } from "./folders.js";

export const SUMMARY_PROPS = [
  "item:Subject",
  "message:From",
  "item:DateTimeReceived",
  "message:IsRead",
  "item:IsDraft",
  "item:HasAttachments",
  "item:Importance",
  "item:Flag",
  "item:Categories",
  "item:Preview",
  "item:DisplayTo",
  "item:ConversationId",
];

export const DETAIL_PROPS = [
  "item:Subject",
  "message:From",
  "message:Sender",
  "message:ToRecipients",
  "message:CcRecipients",
  "message:BccRecipients",
  "message:ReplyTo",
  "item:DateTimeReceived",
  "item:DateTimeSent",
  "message:IsRead",
  "item:IsDraft",
  "item:Attachments",
  "item:Importance",
  "item:Flag",
  "item:Categories",
  "item:ConversationId",
  "message:InternetMessageId",
  "item:ParentFolderId",
  "item:ItemClass",
  "item:Body",
];

export type ViewFilter = "all" | "unread" | "flagged" | "toOrCcMe";
const VIEW_FILTERS: Record<ViewFilter, string> = { all: "All", unread: "Unread", flagged: "Flagged", toOrCcMe: "ToOrCcMe" };

export type Importance = "low" | "normal" | "high";
export type FlagState = "flagged" | "complete" | "none";

export interface ListMessagesOptions {
  folder?: string;
  limit?: number;
  offset?: number;
  filter?: ViewFilter;
  order?: "newest" | "oldest";
}

export interface SearchOptions {
  text?: string;
  from?: string;
  to?: string;
  subject?: string;
  folders?: string[];
  unread?: boolean;
  flagged?: boolean;
  hasAttachments?: boolean;
  category?: string;
  after?: string;
  before?: string;
  limit?: number;
  scanLimit?: number;
}

export interface ComposeInput {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  bodyFormat?: BodyFormat;
  importance?: Importance;
  attachments?: AttachmentInput[];
}

export interface BatchResult {
  succeeded: string[];
  failed: { id: string; error: string }[];
}

type Summary = ReturnType<typeof messageSummary>;

const PAGE_SIZE = 250;
const BATCH_SIZE = 50;
const SEARCH_SKIP_FOLDERS = new Set(["deleteditems", "junkemail", "outbox", "drafts", "conversationhistory"]);

export class MailService {
  constructor(
    private readonly client: OwaClient,
    private readonly folders: FolderService,
    private readonly attachments: AttachmentService,
  ) {}

  // -------------------------------------------------------------------------
  // Reading

  async list(options: ListMessagesOptions = {}) {
    const folder = options.folder ?? "inbox";
    const folderId = await this.folders.resolve(folder);
    const limit = clamp(options.limit ?? 25, 1, 200);
    const offset = Math.max(0, options.offset ?? 0);
    const page = await this.findPage(folderId, offset, limit, options.filter ?? "all", options.order ?? "newest");
    const nextOffset = page.includesLast ? undefined : offset + page.items.length;
    return {
      folder,
      total: page.total,
      offset,
      count: page.items.length,
      nextOffset,
      messages: page.items.map(messageSummary),
    };
  }

  private async findPage(folderId: string, offset: number, max: number, filter: ViewFilter, order: "newest" | "oldest") {
    const [message] = await this.client.callForMessages("FindItem", {
      ItemShape: itemShape(SUMMARY_PROPS),
      ParentFolderIds: [folderIdRef(folderId)],
      Traversal: "Shallow",
      Paging: indexedPage(offset, max),
      ViewFilter: VIEW_FILTERS[filter],
      SortOrder: sortBy("item:DateTimeReceived", order === "oldest" ? "Ascending" : "Descending"),
    });
    const root = (message?.RootFolder as Json | undefined) ?? {};
    return {
      items: (root.Items as Json[] | undefined) ?? [],
      total: typeof root.TotalItemsInView === "number" ? root.TotalItemsInView : undefined,
      includesLast: root.IncludesLastItemInRange !== false,
    };
  }

  /**
   * Search by scanning the newest messages of each folder and filtering locally.
   * (Server-side content search is not reachable with Azure CLI tokens; see docs/OWA-API.md.)
   */
  async search(options: SearchOptions) {
    const limit = clamp(options.limit ?? 25, 1, 200);
    const scanLimit = clamp(options.scanLimit ?? 1000, 50, 10_000);
    const after = options.after ? parseRangeInput(options.after, "after").getTime() : undefined;
    const before = options.before ? parseRangeInput(options.before, "before").getTime() : undefined;
    const terms = tokenize(options.text);
    const needle = (value: string | undefined) => value?.trim().toLowerCase() || undefined;
    const from = needle(options.from);
    const to = needle(options.to);
    const subject = needle(options.subject);
    const category = needle(options.category);
    const filter: ViewFilter = options.unread ? "unread" : options.flagged ? "flagged" : "all";

    const targets = await this.searchFolders(options.folders);
    const matches: (Summary & { folder: string })[] = [];
    let scanned = 0;
    const incomplete: string[] = [];

    for (const target of targets) {
      let offset = 0;
      for (;;) {
        const page = await this.findPage(target.id, offset, Math.min(PAGE_SIZE, scanLimit - offset), filter, "newest");
        scanned += page.items.length;
        let reachedOlder = false;
        for (const raw of page.items) {
          const summary = messageSummary(raw);
          const received = summary.receivedAt ? Date.parse(summary.receivedAt) : undefined;
          if (after !== undefined && received !== undefined && received < after) {
            reachedOlder = true;
            break;
          }
          if (before !== undefined && received !== undefined && received >= before) continue;
          if (options.unread && summary.isRead !== false) continue;
          if (options.flagged && summary.flag !== "flagged") continue;
          if (options.hasAttachments !== undefined && Boolean(summary.hasAttachments) !== options.hasAttachments) continue;
          if (category && !(summary.categories ?? []).some((c: string) => c.toLowerCase() === category)) continue;
          if (subject && !summary.subject.toLowerCase().includes(subject)) continue;
          if (from && !(summary.from ?? "").toLowerCase().includes(from)) continue;
          if (to && !(summary.to ?? "").toLowerCase().includes(to)) continue;
          if (terms.length) {
            const haystack = [summary.subject, summary.preview, summary.from, summary.to].join(" ").toLowerCase();
            if (!terms.every((term) => haystack.includes(term))) continue;
          }
          matches.push({ ...summary, folder: target.label });
        }
        offset += page.items.length;
        if (reachedOlder || page.includesLast || page.items.length === 0) break;
        if (offset >= scanLimit) {
          incomplete.push(target.label);
          break;
        }
      }
    }

    matches.sort((a, b) => (b.receivedAt ?? "").localeCompare(a.receivedAt ?? ""));
    return {
      totalMatches: matches.length,
      returned: Math.min(limit, matches.length),
      scannedMessages: scanned,
      foldersSearched: targets.map((f) => f.label),
      partiallySearchedFolders: incomplete.length ? incomplete : undefined,
      note: incomplete.length
        ? `Only the newest ${scanLimit} messages of some folders were scanned; raise scanLimit or add an "after" date to search further back.`
        : undefined,
      messages: matches.slice(0, limit),
    };
  }

  private async searchFolders(refs: string[] | undefined): Promise<{ id: string; label: string }[]> {
    if (refs?.length && !(refs.length === 1 && refs[0]!.toLowerCase() === "all")) {
      return Promise.all(refs.map(async (ref) => ({ id: await this.folders.resolve(ref), label: ref })));
    }
    const all = await this.folders.list("mail");
    return all
      .filter((f) => f.wellKnownName !== "msgfolderroot" && !SEARCH_SKIP_FOLDERS.has(f.wellKnownName ?? "") && (f.totalCount ?? 0) > 0)
      .map((f) => ({ id: f.id, label: f.path }));
  }

  async get(id: string, options: { format?: BodyFormat; maxBodyChars?: number; markAsRead?: boolean } = {}) {
    const format = options.format ?? "text";
    const [message] = await this.client.callForMessages("GetItem", {
      ItemShape: itemShape(DETAIL_PROPS, format === "html" ? { BodyType: "HTML", FilterHtmlContent: true } : { BodyType: "Text" }),
      ItemIds: [itemId(id)],
    });
    const raw = (message?.Items as Json[] | undefined)?.[0];
    if (!raw) throw new OutlookError("Message not found");
    const detail = messageDetail(raw, { maxChars: options.maxBodyChars });
    if (options.markAsRead && raw.IsRead === false) {
      await this.markRead([id], true);
      detail.isRead = true;
    }
    return detail;
  }

  async thread(selector: { messageId?: string; conversationId?: string }, options: { maxBodyChars?: number } = {}) {
    let conversationId = selector.conversationId;
    if (!conversationId) {
      if (!selector.messageId) throw new OutlookError("Pass messageId or conversationId");
      const [message] = await this.client.callForMessages("GetItem", {
        ItemShape: itemShape(["item:ConversationId"]),
        ItemIds: [itemId(selector.messageId)],
      });
      conversationId = (((message?.Items as Json[] | undefined)?.[0]?.ConversationId as Json | undefined)?.Id as string) ?? undefined;
      if (!conversationId) throw new OutlookError("This message has no conversation ID");
    }
    const [message] = await this.client.callForMessages("GetConversationItems", {
      Conversations: [{ __type: t("ConversationRequestType"), ConversationId: itemId(conversationId) }],
      ItemShape: itemShape(
        ["item:Subject", "message:From", "message:ToRecipients", "message:CcRecipients", "item:DateTimeReceived", "item:DateTimeSent",
          "message:IsRead", "item:Attachments", "item:ParentFolderId", "item:UniqueBody"],
        { BodyType: "Text" },
      ),
      ShapeName: "ItemPart",
      SortOrder: "DateOrderDescending",
      MaxItemsToReturn: 100,
    });
    const nodes = ((message?.Conversation as Json | undefined)?.ConversationNodes as Json[] | undefined) ?? [];
    const items = nodes.flatMap((node) => (node.Items as Json[] | undefined) ?? []);
    const messages = items
      .map((raw) => messageDetail(raw, { maxChars: options.maxBodyChars ?? 4000 }))
      .sort((a, b) => (a.receivedAt ?? a.sentAt ?? "").localeCompare(b.receivedAt ?? b.sentAt ?? ""));
    return { conversationId, count: messages.length, messages };
  }

  // -------------------------------------------------------------------------
  // Composing

  private messageItem(input: ComposeInput): Json {
    const item: Json = { __type: t("Message") };
    if (input.subject !== undefined) item.Subject = input.subject;
    if (input.body !== undefined) item.Body = bodyContent(input.body, input.bodyFormat);
    if (input.to?.length) item.ToRecipients = mailRecipients(input.to);
    if (input.cc?.length) item.CcRecipients = mailRecipients(input.cc);
    if (input.bcc?.length) item.BccRecipients = mailRecipients(input.bcc);
    if (input.importance) item.Importance = capitalize(input.importance);
    return item;
  }

  /** Send a new message. With attachments: draft → attach → send. */
  async send(input: ComposeInput): Promise<{ sent: true; attachments?: number }> {
    if (!input.to?.length && !input.cc?.length && !input.bcc?.length) {
      throw new OutlookError("At least one recipient (to, cc or bcc) is required");
    }
    if (input.attachments?.length) {
      const { id } = await this.createDraft(input);
      await this.sendDraft(id);
      return { sent: true, attachments: input.attachments.length };
    }
    const item = this.messageItem(input);
    await this.client.callForMessages("CreateItem", {
      MessageDisposition: "SendAndSaveCopy",
      SavedItemFolderId: targetFolder("sentitems"),
      Items: [item],
    });
    return { sent: true };
  }

  async createDraft(input: ComposeInput): Promise<{ id: string; attachments?: { id: string; name?: string }[] }> {
    const [message] = await this.client.callForMessages("CreateItem", {
      MessageDisposition: "SaveOnly",
      SavedItemFolderId: targetFolder("drafts"),
      Items: [this.messageItem(input)],
    });
    const id = createdItemId(message);
    const added = await this.addAttachments(id, input.attachments);
    return { id, attachments: added.length ? added : undefined };
  }

  async updateDraft(id: string, input: Omit<ComposeInput, "attachments">): Promise<{ id: string }> {
    const updates: Json[] = [];
    const set = (uri: string, values: Json) => updates.push(setItemField(propertyUri(uri), "Message", values));
    if (input.subject !== undefined) set("item:Subject", { Subject: input.subject });
    if (input.body !== undefined) set("item:Body", { Body: bodyContent(input.body, input.bodyFormat) });
    if (input.importance) set("item:Importance", { Importance: capitalize(input.importance) });
    for (const [field, uri, list] of [
      ["ToRecipients", "message:ToRecipients", input.to],
      ["CcRecipients", "message:CcRecipients", input.cc],
      ["BccRecipients", "message:BccRecipients", input.bcc],
    ] as const) {
      if (list === undefined) continue;
      if (list.length) set(uri, { [field]: mailRecipients(list) });
      else updates.push(deleteItemField(propertyUri(uri)));
    }
    if (!updates.length) throw new OutlookError("Nothing to update");
    await this.client.callForMessages("UpdateItem", {
      ConflictResolution: "AlwaysOverwrite",
      MessageDisposition: "SaveOnly",
      SuppressReadReceipts: true,
      ItemChanges: [itemChange(id, updates)],
    });
    return { id };
  }

  /**
   * Send an existing draft. OWA has no SendItem action; sending is an UpdateItem
   * with MessageDisposition SendAndSaveCopy, which needs at least one change, so
   * the current subject is written back unchanged.
   */
  async sendDraft(id: string): Promise<{ sent: true }> {
    const [message] = await this.client.callForMessages("GetItem", {
      ItemShape: itemShape(["item:Subject", "item:IsDraft"]),
      ItemIds: [itemId(id)],
    });
    const draft = (message?.Items as Json[] | undefined)?.[0];
    if (!draft) throw new OutlookError("Draft not found");
    if (draft.IsDraft === false) throw new OutlookError("This item is not a draft (it has already been sent)");
    await this.client.callForMessages("UpdateItem", {
      ConflictResolution: "AlwaysOverwrite",
      MessageDisposition: "SendAndSaveCopy",
      SendMeetingInvitations: "SendToNone",
      SuppressReadReceipts: true,
      ItemChanges: [itemChange(id, [setItemField(propertyUri("item:Subject"), "Message", { Subject: draft.Subject ?? "" })])],
    });
    return { sent: true };
  }

  /** Reply or reply-all. The server quotes the original message below the new body. */
  async reply(
    id: string,
    input: { body: string; bodyFormat?: BodyFormat; replyAll?: boolean; send: boolean; attachments?: AttachmentInput[] },
  ) {
    const item: Json = {
      __type: t(input.replyAll ? "ReplyAllToItem" : "ReplyToItem"),
      ReferenceItemId: itemId(id),
      ShouldIgnoreChangeKey: true,
      NewBodyContent: bodyContent(input.body, input.bodyFormat),
    };
    return this.respond(item, input.send, input.attachments);
  }

  async forward(
    id: string,
    input: { to: string[]; cc?: string[]; bcc?: string[]; body?: string; bodyFormat?: BodyFormat; send: boolean; attachments?: AttachmentInput[] },
  ) {
    if (!input.to.length) throw new OutlookError("Forwarding needs at least one recipient");
    const item: Json = {
      __type: t("ForwardItem"),
      ReferenceItemId: itemId(id),
      ShouldIgnoreChangeKey: true,
      ToRecipients: mailRecipients(input.to),
    };
    if (input.cc?.length) item.CcRecipients = mailRecipients(input.cc);
    if (input.bcc?.length) item.BccRecipients = mailRecipients(input.bcc);
    item.NewBodyContent = bodyContent(input.body ?? "", input.bodyFormat);
    return this.respond(item, input.send, input.attachments);
  }

  private async respond(item: Json, send: boolean, attachments: AttachmentInput[] | undefined) {
    if (send && !attachments?.length) {
      await this.client.callForMessages("CreateItem", {
        MessageDisposition: "SendAndSaveCopy",
        SavedItemFolderId: targetFolder("sentitems"),
        Items: [item],
      });
      return { sent: true as const };
    }
    const [message] = await this.client.callForMessages("CreateItem", {
      MessageDisposition: "SaveOnly",
      SavedItemFolderId: targetFolder("drafts"),
      Items: [item],
    });
    const draftId = createdItemId(message);
    await this.addAttachments(draftId, attachments);
    if (send) {
      await this.sendDraft(draftId);
      return { sent: true as const };
    }
    return { sent: false as const, draftId };
  }

  private async addAttachments(id: string, attachments: AttachmentInput[] | undefined) {
    const added: { id: string; name?: string }[] = [];
    for (const attachment of attachments ?? []) {
      const info = await this.attachments.add(id, attachment);
      added.push({ id: info.id, name: info.name });
    }
    return added;
  }

  // -------------------------------------------------------------------------
  // Triage (batched)

  async markRead(ids: string[], read: boolean): Promise<BatchResult> {
    return this.updateEach(ids, () => [setItemField(propertyUri("IsRead"), "Message", { IsRead: read })]);
  }

  async setFlag(ids: string[], state: FlagState): Promise<BatchResult> {
    const status = state === "flagged" ? "Flagged" : state === "complete" ? "Complete" : "NotFlagged";
    return this.updateEach(ids, () => [
      setItemField(propertyUri("item:Flag"), "Message", { Flag: { __type: t("FlagType"), FlagStatus: status } }),
    ]);
  }

  async setCategories(ids: string[], categories: string[], mode: "set" | "add" | "remove"): Promise<BatchResult> {
    const wanted = unique(categories.map((c) => c.trim()).filter(Boolean));
    let current = new Map<string, string[]>();
    if (mode !== "set") current = await this.currentCategories(ids);
    return this.updateEach(ids, (id) => {
      const existing = current.get(id) ?? [];
      const next =
        mode === "set"
          ? wanted
          : mode === "add"
            ? unique([...existing, ...wanted])
            : existing.filter((c) => !wanted.some((w) => w.toLowerCase() === c.toLowerCase()));
      return next.length
        ? [setItemField(propertyUri("item:Categories"), "Message", { Categories: next })]
        : [deleteItemField(propertyUri("item:Categories"))];
    });
  }

  private async currentCategories(ids: string[]): Promise<Map<string, string[]>> {
    const out = new Map<string, string[]>();
    for (const group of chunk(ids, BATCH_SIZE)) {
      const result = await this.client.call("GetItem", {
        ItemShape: itemShape(["item:Categories"]),
        ItemIds: group.map((id) => itemId(id)),
      });
      responseMessages(result).forEach((message, index) => {
        const item = (message.Items as Json[] | undefined)?.[0];
        if (item) out.set(group[index]!, (item.Categories as string[] | undefined) ?? []);
      });
    }
    return out;
  }

  private async updateEach(ids: string[], updatesFor: (id: string) => Json[]): Promise<BatchResult> {
    return this.batch("UpdateItem", ids, (group) => ({
      ConflictResolution: "AlwaysOverwrite",
      MessageDisposition: "SaveOnly",
      SuppressReadReceipts: true,
      ItemChanges: group.map((id) => itemChange(id, updatesFor(id))),
    }));
  }

  async move(ids: string[], destination: string): Promise<BatchResult & { destination: string }> {
    const folderId = await this.folders.resolve(destination);
    const result = await this.batch("MoveItem", ids, (group) => ({
      ToFolderId: targetFolder(folderId),
      ItemIds: group.map((id) => itemId(id)),
      ReturnNewItemIds: true,
    }));
    return { ...result, destination };
  }

  async copy(ids: string[], destination: string): Promise<BatchResult & { destination: string; newIds: string[] }> {
    const folderId = await this.folders.resolve(destination);
    const newIds: string[] = [];
    const result = await this.batch(
      "CopyItem",
      ids,
      (group) => ({ ToFolderId: targetFolder(folderId), ItemIds: group.map((id) => itemId(id)), ReturnNewItemIds: true }),
      (message) => {
        const id = ((message.Items as Json[] | undefined)?.[0]?.ItemId as Json | undefined)?.Id;
        if (typeof id === "string") newIds.push(id);
      },
    );
    return { ...result, destination, newIds };
  }

  async delete(ids: string[], permanent: boolean): Promise<BatchResult> {
    return this.batch("DeleteItem", ids, (group) => ({
      ItemIds: group.map((id) => itemId(id)),
      DeleteType: permanent ? "HardDelete" : "MoveToDeletedItems",
      SendMeetingCancellations: "SendToNone",
    }));
  }

  /** Run an action over ids in chunks and report per-item outcomes (responses are in request order). */
  private async batch(action: string, ids: string[], build: (group: string[]) => Json, onSuccess?: (message: Json) => void): Promise<BatchResult> {
    const list = unique(ids);
    if (!list.length) throw new OutlookError("No message IDs given");
    const succeeded: string[] = [];
    const failed: { id: string; error: string }[] = [];
    for (const group of chunk(list, BATCH_SIZE)) {
      let messages: Json[];
      try {
        messages = responseMessages(await this.client.call(action, build(group)));
      } catch (error) {
        for (const id of group) failed.push({ id, error: errorMessage(error) });
        continue;
      }
      group.forEach((id, index) => {
        const message = messages[index];
        if (!message) failed.push({ id, error: "No response from server" });
        else if (isErrorMessage(message)) {
          failed.push({ id, error: describeResponseCode(action, String(message.ResponseCode), message.MessageText as string | undefined) });
        } else {
          succeeded.push(id);
          onSuccess?.(message);
        }
      });
    }
    return { succeeded, failed };
  }
}

function createdItemId(message: Json | undefined): string {
  assertSuccess("CreateItem", message);
  const id = (((message?.Items as Json[] | undefined)?.[0]?.ItemId as Json | undefined)?.Id as string | undefined) ?? undefined;
  if (!id) throw new OutlookError("The server did not return the new item's ID");
  return id;
}

function tokenize(text: string | undefined): string[] {
  if (!text) return [];
  const terms: string[] = [];
  for (const match of text.toLowerCase().matchAll(/"([^"]+)"|(\S+)/g)) {
    const term = (match[1] ?? match[2] ?? "").trim();
    if (term) terms.push(term);
  }
  return terms;
}

function unique<T>(items: T[]): T[] {
  return [...new Set(items)];
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, Math.floor(value)));
}

function capitalize(text: string): string {
  return text.charAt(0).toUpperCase() + text.slice(1);
}
