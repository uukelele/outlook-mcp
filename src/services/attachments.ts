import { readFile, stat } from "node:fs/promises";
import { basename, extname } from "node:path";
import { OutlookError } from "../errors.js";
import { attachmentId, itemId, itemShape, t, type Json } from "../owa/builders.js";
import type { OwaClient } from "../owa/client.js";
import { attachmentSummary } from "../owa/normalize.js";
import type { MailboxService } from "./mailbox.js";

export interface AttachmentInput {
  /** File name shown to recipients. Defaults to the basename of `path`. */
  name?: string;
  /** Local file to attach. */
  path?: string;
  /** Base64-encoded content. */
  contentBase64?: string;
  /** Plain text content (UTF-8). */
  text?: string;
  contentType?: string;
  isInline?: boolean;
  contentId?: string;
}

export interface AttachmentInfo {
  id: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  kind?: string;
}

export interface DownloadedAttachment extends AttachmentInfo {
  data: Buffer;
}

const CONTENT_TYPES: Record<string, string> = {
  ".txt": "text/plain",
  ".md": "text/markdown",
  ".csv": "text/csv",
  ".html": "text/html",
  ".htm": "text/html",
  ".json": "application/json",
  ".xml": "application/xml",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".svg": "image/svg+xml",
  ".ics": "text/calendar",
  ".zip": "application/zip",
  ".doc": "application/msword",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xls": "application/vnd.ms-excel",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".ppt": "application/vnd.ms-powerpoint",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

export function guessContentType(name: string): string {
  return CONTENT_TYPES[extname(name).toLowerCase()] ?? "application/octet-stream";
}

export class AttachmentService {
  constructor(
    private readonly client: OwaClient,
    private readonly mailbox: MailboxService,
    private readonly maxBytes: number,
  ) {}

  async list(id: string): Promise<AttachmentInfo[]> {
    const [message] = await this.client.callForMessages("GetItem", {
      ItemShape: itemShape(["item:Attachments"]),
      ItemIds: [itemId(id)],
    });
    const item = (message?.Items as Json[] | undefined)?.[0];
    return ((item?.Attachments as Json[] | undefined) ?? []).map((a) => attachmentSummary(a) as AttachmentInfo);
  }

  /** Download a file attachment of an item, selected by attachment ID or file name. */
  async download(parentId: string, selector: { attachmentId?: string; name?: string }): Promise<DownloadedAttachment> {
    const attachments = await this.list(parentId);
    if (!attachments.length) throw new OutlookError("This item has no attachments");
    let match: AttachmentInfo | undefined;
    if (selector.attachmentId) {
      match = attachments.find((a) => a.id === selector.attachmentId);
    } else if (selector.name) {
      const wanted = selector.name.toLowerCase();
      const byName = attachments.filter((a) => a.name?.toLowerCase() === wanted);
      if (byName.length > 1) throw new OutlookError(`Several attachments are named "${selector.name}"; pass attachmentId instead`);
      match = byName[0];
    } else if (attachments.length === 1) {
      match = attachments[0];
    } else {
      throw new OutlookError(`This item has ${attachments.length} attachments; pass attachmentId or name`);
    }
    if (!match) {
      throw new OutlookError(
        `Attachment not found. Available: ${attachments.map((a) => `"${a.name}" (${a.id})`).join(", ")}`,
      );
    }
    if (match.kind === "item") {
      throw new OutlookError(`"${match.name}" is an attached Outlook item (e.g. an email), not a file; it cannot be downloaded as bytes`);
    }
    if (match.kind === "link") {
      throw new OutlookError(`"${match.name}" is a cloud link attachment (OneDrive/SharePoint); open it from the original message`);
    }
    if (match.size !== undefined && match.size > this.maxBytes) {
      throw new OutlookError(`"${match.name}" is ${match.size} bytes, above the ${this.maxBytes}-byte limit (OUTLOOK_MAX_DOWNLOAD_MB)`);
    }
    const { data, contentType } = await this.fetchBytes(match.id);
    return { ...match, contentType: match.contentType ?? contentType, data };
  }

  /** Two-step download used by Outlook on the web: short-lived token, then GET from the attachments host. */
  private async fetchBytes(id: string): Promise<{ data: Buffer; contentType?: string }> {
    const token = await this.client.call<unknown>(
      "GetAttachmentDownloadToken",
      { __type: t("GetAttachmentDownloadTokenRequest"), UseNewDownloadTokenLifetime: true },
      { raw: true },
    );
    if (typeof token !== "string" || !token) {
      throw new OutlookError("GetAttachmentDownloadToken did not return a token");
    }
    const mailbox = await this.mailbox.emailAddress();
    if (!mailbox) throw new OutlookError("Cannot determine the mailbox address needed for attachment downloads");
    const url = `${this.client.attachmentsOrigin}/owa/${encodeURIComponent(mailbox)}/service.svc/s/GetFileAttachment?id=${encodeURIComponent(id)}`;
    return this.client.download(url, { "x-token": token }, this.maxBytes);
  }

  async add(parentId: string, input: AttachmentInput): Promise<AttachmentInfo> {
    const { name, content, contentType } = await this.readInput(input);
    const attachment: Json = {
      __type: t("FileAttachment"),
      Name: name,
      ContentType: contentType,
      IsInline: Boolean(input.isInline),
      Content: content.toString("base64"),
    };
    if (input.contentId) attachment.ContentId = input.contentId;
    const [message] = await this.client.callForMessages("CreateAttachment", {
      ParentItemId: itemId(parentId),
      Attachments: [attachment],
      RequireImageType: false,
      IncludeContentIdInResponse: Boolean(input.contentId),
      ClientSupportsIrm: true,
    });
    const created = (message?.Attachments as Json[] | undefined)?.[0];
    const id = (created?.AttachmentId as Json | undefined)?.Id as string | undefined;
    if (!id) throw new OutlookError("CreateAttachment succeeded but returned no attachment ID");
    return { id, name, contentType, size: content.length, isInline: input.isInline || undefined };
  }

  async remove(id: string): Promise<void> {
    await this.client.callForMessages("DeleteAttachment", { AttachmentIds: [attachmentId(id)] });
  }

  private async readInput(input: AttachmentInput): Promise<{ name: string; content: Buffer; contentType: string }> {
    const sources = [input.path, input.contentBase64, input.text].filter((v) => v !== undefined);
    if (sources.length !== 1) {
      throw new OutlookError("Each attachment needs exactly one of: path, contentBase64, text");
    }
    let content: Buffer;
    let name = input.name;
    if (input.path !== undefined) {
      let size: number;
      try {
        size = (await stat(input.path)).size;
      } catch (error) {
        throw new OutlookError(`Cannot read attachment file ${input.path}: ${(error as Error).message}`);
      }
      if (size > this.maxBytes) throw new OutlookError(`${input.path} is ${size} bytes, above the ${this.maxBytes}-byte limit`);
      content = await readFile(input.path);
      name ??= basename(input.path);
    } else if (input.contentBase64 !== undefined) {
      const cleaned = input.contentBase64.replace(/^data:[^;]+;base64,/, "").replace(/\s+/g, "");
      if (!/^[A-Za-z0-9+/_-]*={0,2}$/.test(cleaned)) throw new OutlookError("contentBase64 is not valid base64");
      content = Buffer.from(cleaned, "base64");
    } else {
      content = Buffer.from(input.text!, "utf8");
    }
    if (!name) throw new OutlookError("Attachment name is required when not attaching a file path");
    if (content.length > this.maxBytes) throw new OutlookError(`Attachment "${name}" is ${content.length} bytes, above the ${this.maxBytes}-byte limit`);
    return { name, content, contentType: input.contentType ?? guessContentType(name) };
  }
}
