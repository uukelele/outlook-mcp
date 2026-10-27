/**
 * Builders for OWA service.svc JSON payloads.
 *
 * The service deserializes with .NET DataContractJsonSerializer, which requires
 * `__type` to be the first key of every polymorphic object. Every builder here
 * therefore writes `__type` first; JSON.stringify preserves insertion order.
 */

export type Json = Record<string, unknown>;

export const DEFAULT_SERVER_VERSION = "V2018_01_08";

export function t(name: string): string {
  return `${name}:#Exchange`;
}

export interface EnvelopeOptions {
  /** Type name used for `<X>JsonRequest`/`<X>Request` when it differs from the action (e.g. CreateCalendarEvent → CreateItem). */
  jsonType?: string;
  version?: string;
  /** Windows time zone ID for TimeZoneContext. */
  timeZone?: string;
}

export function requestHeader(version = DEFAULT_SERVER_VERSION, timeZone = "UTC"): Json {
  return {
    __type: t("JsonRequestHeaders"),
    RequestServerVersion: version,
    TimeZoneContext: {
      __type: t("TimeZoneContext"),
      TimeZoneDefinition: { __type: t("TimeZoneDefinitionType"), Id: timeZone },
    },
  };
}

export function envelope(action: string, body: Json, options: EnvelopeOptions = {}): Json {
  const type = options.jsonType ?? action;
  const { __type: _ignored, ...rest } = body;
  return {
    __type: t(`${type}JsonRequest`),
    Header: requestHeader(options.version, options.timeZone),
    Body: { __type: t(`${type}Request`), ...rest },
  };
}

// ---------------------------------------------------------------------------
// Folders

/** Distinguished folder IDs that are safe to pass straight to the service. */
export const DISTINGUISHED_FOLDERS = new Set([
  "msgfolderroot",
  "root",
  "inbox",
  "drafts",
  "sentitems",
  "deleteditems",
  "junkemail",
  "outbox",
  "archive",
  "calendar",
  "contacts",
  "directory",
  "tasks",
  "notes",
  "journal",
  "conversationhistory",
  "searchfolders",
  "recoverableitemsdeletions",
]);

const FOLDER_ALIASES: Record<string, string> = {
  sent: "sentitems",
  "sent items": "sentitems",
  deleted: "deleteditems",
  "deleted items": "deleteditems",
  trash: "deleteditems",
  bin: "deleteditems",
  junk: "junkemail",
  "junk email": "junkemail",
  spam: "junkemail",
  draft: "drafts",
  "mail root": "msgfolderroot",
  "top of information store": "msgfolderroot",
};

/** Map a user-facing folder name to a distinguished folder ID, if it is one. */
export function distinguishedFolder(ref: string): string | undefined {
  const key = ref.trim().toLowerCase();
  const name = FOLDER_ALIASES[key] ?? key;
  return DISTINGUISHED_FOLDERS.has(name) ? name : undefined;
}

/** Exchange item and folder IDs are long base64 strings; names and paths are not. */
export function looksLikeExchangeId(ref: string): boolean {
  return ref.length >= 40 && /^[A-Za-z0-9+/_=-]+$/.test(ref);
}

export function folderIdRef(ref: string): Json {
  const distinguished = distinguishedFolder(ref);
  if (distinguished) return { __type: t("DistinguishedFolderId"), Id: distinguished };
  return { __type: t("FolderId"), Id: ref };
}

export function targetFolder(ref: string): Json {
  return { __type: t("TargetFolderId"), BaseFolderId: folderIdRef(ref) };
}

// ---------------------------------------------------------------------------
// Items and properties

export function itemId(id: string, changeKey?: string): Json {
  return changeKey ? { __type: t("ItemId"), Id: id, ChangeKey: changeKey } : { __type: t("ItemId"), Id: id };
}

export function attachmentId(id: string): Json {
  return { __type: t("AttachmentId"), Id: id };
}

export function propertyUri(fieldUri: string): Json {
  return { __type: t("PropertyUri"), FieldURI: fieldUri };
}

export function indexedPropertyUri(fieldUri: string, fieldIndex: string): Json {
  return { __type: t("DictionaryPropertyUri"), FieldURI: fieldUri, FieldIndex: fieldIndex };
}

export type BodyFormat = "text" | "html";

export function bodyContent(content: string, format: BodyFormat = "text"): Json {
  return { __type: t("BodyContentType"), BodyType: format === "html" ? "HTML" : "Text", Value: content };
}

export function itemShape(additional: string[], extra: Json = {}): Json {
  return {
    __type: t("ItemResponseShape"),
    BaseShape: "IdOnly",
    ...extra,
    AdditionalProperties: additional.map(propertyUri),
  };
}

export function indexedPage(offset: number, max: number): Json {
  return { __type: t("IndexedPageView"), BasePoint: "Beginning", Offset: offset, MaxEntriesReturned: max };
}

export function sortBy(fieldUri: string, order: "Ascending" | "Descending" = "Descending"): Json[] {
  return [{ __type: t("SortResults"), Order: order, Path: propertyUri(fieldUri) }];
}

export function setItemField(path: Json, itemType: string, values: Json): Json {
  return { __type: t("SetItemField"), Path: path, Item: { __type: t(itemType), ...values } };
}

export function deleteItemField(path: Json): Json {
  return { __type: t("DeleteItemField"), Path: path };
}

export function itemChange(id: string, updates: Json[]): Json {
  return { __type: t("ItemChange"), ItemId: itemId(id), Updates: updates };
}

// ---------------------------------------------------------------------------
// Addresses

export interface Address {
  email: string;
  name?: string;
}

const EMAIL_PATTERN = /^[^\s@<>]+@[^\s@<>]+$/;

/** Parse `"Name <user@example.com>"` or `user@example.com`. */
export function parseAddress(input: string): Address {
  const text = input.trim();
  const angle = /^(.*?)<\s*([^<>\s]+)\s*>$/.exec(text);
  const email = (angle ? angle[2] : text)!.trim();
  if (!EMAIL_PATTERN.test(email)) {
    throw new Error(`"${input}" is not a valid email address`);
  }
  const name = angle?.[1]?.trim().replace(/^"(.*)"$/, "$1").trim();
  return name ? { email, name } : { email };
}

/** Mail recipients in OWA JSON are flat objects without __type (verified: nested Mailbox shapes do not resolve). */
export function mailRecipients(addresses: string[] | undefined): Json[] {
  return (addresses ?? []).map((input) => {
    const { email, name } = parseAddress(input);
    return name ? { Name: name, EmailAddress: email, RoutingType: "SMTP" } : { EmailAddress: email, RoutingType: "SMTP" };
  });
}

/** Calendar attendees are wrapped in a Mailbox object. */
export function attendees(addresses: string[] | undefined): Json[] {
  return (addresses ?? []).map((input) => {
    const { email, name } = parseAddress(input);
    const mailbox: Json = name ? { Name: name, EmailAddress: email, RoutingType: "SMTP" } : { EmailAddress: email, RoutingType: "SMTP" };
    return { __type: t("AttendeeType"), Mailbox: mailbox };
  });
}

// ---------------------------------------------------------------------------
// Utilities

export function chunk<T>(items: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
  return out;
}

/** True if `__type`, when present, is the first key of every object in the tree. */
export function typeKeysFirst(value: unknown): boolean {
  if (Array.isArray(value)) return value.every(typeKeysFirst);
  if (value && typeof value === "object") {
    const keys = Object.keys(value);
    if (keys.includes("__type") && keys[0] !== "__type") return false;
    return Object.values(value).every(typeKeysFirst);
  }
  return true;
}
