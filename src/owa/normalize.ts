/** Convert raw OWA objects into compact, predictable JSON for tool output. */

type Raw = Record<string, any>;

export interface Person {
  name?: string;
  email?: string;
}

/** Remove undefined/null/empty-string/empty-array values (recursively). */
export function compact<T>(value: T): T {
  if (Array.isArray(value)) {
    return value.map((entry) => compact(entry)).filter((entry) => !isEmpty(entry)) as T;
  }
  if (value && typeof value === "object" && !(value instanceof Date)) {
    const out: Raw = {};
    for (const [key, entry] of Object.entries(value)) {
      const cleaned = compact(entry);
      if (!isEmpty(cleaned)) out[key] = cleaned;
    }
    return out as T;
  }
  return value;
}

function isEmpty(value: unknown): boolean {
  if (value === undefined || value === null || value === "") return true;
  if (Array.isArray(value)) return value.length === 0;
  if (typeof value === "object" && !(value instanceof Date)) return Object.keys(value as object).length === 0;
  return false;
}

function str(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value : undefined;
}

/** Accepts `{Mailbox: {...}}` or a flat mailbox; prefers the SMTP address over Exchange legacy DNs. */
export function person(raw: Raw | undefined | null): Person | undefined {
  if (!raw) return undefined;
  const mailbox: Raw = raw.Mailbox ?? raw;
  const address = str(mailbox.SmtpEmailAddress) ?? (mailbox.RoutingType === "EX" ? undefined : str(mailbox.EmailAddress));
  const name = str(mailbox.Name);
  if (!address && !name) return undefined;
  return { name: name && name !== address ? name : undefined, email: address };
}

export function people(list: Raw[] | undefined | null): Person[] {
  return (list ?? []).map((entry) => person(entry)).filter((entry): entry is Person => Boolean(entry));
}

export function formatPerson(p: Person | undefined): string | undefined {
  if (!p) return undefined;
  if (p.name && p.email) return `${p.name} <${p.email}>`;
  return p.email ?? p.name;
}

export function flagStatus(raw: Raw | undefined): "flagged" | "complete" | undefined {
  switch (raw?.FlagStatus) {
    case "Flagged":
      return "flagged";
    case "Complete":
      return "complete";
    default:
      return undefined;
  }
}

function importance(value: unknown): string | undefined {
  return value === "High" || value === "Low" ? String(value).toLowerCase() : undefined;
}

export interface BodyOptions {
  maxChars?: number;
}

/** Body as flat fields: `body` (string), `bodyFormat`, and `bodyTruncated` when cut short. */
export function bodyFields(
  raw: Raw | undefined,
  options: BodyOptions = {},
): { body?: string; bodyFormat?: "text" | "html"; bodyTruncated?: boolean } {
  if (!raw || typeof raw.Value !== "string") return {};
  const bodyFormat = raw.BodyType === "HTML" ? "html" : "text";
  let body: string = raw.Value.replace(/\r\n/g, "\n");
  if (bodyFormat === "text") body = body.replace(/\n{3,}/g, "\n\n").trim();
  const max = options.maxChars;
  if (max && body.length > max) {
    return { body: `${body.slice(0, max)}\n…[truncated ${body.length - max} characters]`, bodyFormat, bodyTruncated: true };
  }
  return { body, bodyFormat };
}

export function attachmentSummary(raw: Raw) {
  const kind = String(raw.__type ?? "").replace(/:#Exchange$/, "");
  return compact({
    id: raw.AttachmentId?.Id as string | undefined,
    name: str(raw.Name),
    contentType: str(raw.ContentType),
    size: typeof raw.Size === "number" ? raw.Size : undefined,
    isInline: raw.IsInline === true ? true : undefined,
    kind: kind === "FileAttachment" ? undefined : kind === "ItemAttachment" ? "item" : kind === "ReferenceAttachment" ? "link" : kind || undefined,
  });
}

/** File attachments, with inline (signature/embedded) images reduced to a count. */
function attachmentFields(list: Raw[] | undefined) {
  const all = (list ?? []).map(attachmentSummary);
  const inline = all.filter((a) => a.isInline);
  return {
    attachments: all.filter((a) => !a.isInline),
    inlineImages: inline.length || undefined,
  };
}

export function messageSummary(raw: Raw) {
  return compact({
    id: raw.ItemId?.Id as string,
    conversationId: raw.ConversationId?.Id as string | undefined,
    subject: str(raw.Subject) ?? "(no subject)",
    from: formatPerson(person(raw.From)),
    to: str(raw.DisplayTo),
    receivedAt: str(raw.DateTimeReceived),
    isRead: typeof raw.IsRead === "boolean" ? raw.IsRead : undefined,
    isDraft: raw.IsDraft === true ? true : undefined,
    hasAttachments: raw.HasAttachments === true ? true : undefined,
    importance: importance(raw.Importance),
    flag: flagStatus(raw.Flag),
    categories: Array.isArray(raw.Categories) ? raw.Categories : undefined,
    preview: str(raw.Preview)?.replace(/\s+/g, " ").trim(),
  });
}

export function messageDetail(raw: Raw, options: BodyOptions = {}) {
  return compact({
    id: raw.ItemId?.Id as string,
    conversationId: raw.ConversationId?.Id as string | undefined,
    folderId: raw.ParentFolderId?.Id as string | undefined,
    subject: str(raw.Subject) ?? "(no subject)",
    from: person(raw.From),
    sender: sameMailbox(raw.Sender, raw.From) ? undefined : person(raw.Sender),
    to: people(raw.ToRecipients),
    cc: people(raw.CcRecipients),
    bcc: people(raw.BccRecipients),
    replyTo: people(raw.ReplyTo),
    receivedAt: str(raw.DateTimeReceived),
    sentAt: str(raw.DateTimeSent),
    isRead: typeof raw.IsRead === "boolean" ? raw.IsRead : undefined,
    isDraft: raw.IsDraft === true ? true : undefined,
    importance: importance(raw.Importance),
    flag: flagStatus(raw.Flag),
    categories: Array.isArray(raw.Categories) ? raw.Categories : undefined,
    internetMessageId: str(raw.InternetMessageId),
    itemClass: raw.ItemClass && raw.ItemClass !== "IPM.Note" ? raw.ItemClass : undefined,
    ...attachmentFields(raw.Attachments),
    ...bodyFields(raw.UniqueBody ?? raw.Body, options),
  });
}

function sameMailbox(a: Raw | undefined, b: Raw | undefined): boolean {
  if (!a) return true;
  return person(a)?.email?.toLowerCase() === person(b)?.email?.toLowerCase();
}

export function folder(raw: Raw, path?: string) {
  return compact({
    id: raw.FolderId?.Id as string,
    name: str(raw.DisplayName),
    path,
    wellKnownName: str(raw.DistinguishedFolderId),
    parentId: raw.ParentFolderId?.Id as string | undefined,
    folderClass: str(raw.FolderClass),
    totalCount: typeof raw.TotalCount === "number" ? raw.TotalCount : undefined,
    unreadCount: typeof raw.UnreadCount === "number" ? raw.UnreadCount : undefined,
    childFolderCount: typeof raw.ChildFolderCount === "number" && raw.ChildFolderCount > 0 ? raw.ChildFolderCount : undefined,
  });
}

function locationName(raw: Raw): string | undefined {
  if (typeof raw.Location === "string") return str(raw.Location);
  return str(raw.Location?.DisplayName) ?? str(raw.Locations?.[0]?.DisplayName);
}

const RESPONSE_TYPES: Record<string, string> = {
  Organizer: "organizer",
  Accept: "accepted",
  Tentative: "tentative",
  Decline: "declined",
  NoResponseReceived: "notResponded",
  Unknown: "unknown",
};

function responseType(value: unknown): string | undefined {
  return typeof value === "string" ? (RESPONSE_TYPES[value] ?? value) : undefined;
}

function attendee(raw: Raw, type: "required" | "optional" | "resource") {
  return compact({ ...person(raw), type, response: responseType(raw.ResponseType) });
}

export function eventSummary(raw: Raw) {
  return compact({
    id: raw.ItemId?.Id as string,
    subject: str(raw.Subject) ?? "(no subject)",
    start: str(raw.Start),
    end: str(raw.End),
    isAllDay: raw.IsAllDayEvent === true ? true : undefined,
    location: locationName(raw),
    organizer: formatPerson(person(raw.Organizer)),
    myResponse: raw.IsMeeting ? responseType(raw.ResponseType) : undefined,
    showAs: str(raw.FreeBusyType),
    isMeeting: raw.IsMeeting === true ? true : undefined,
    isRecurring: raw.IsRecurring === true ? true : undefined,
    isCancelled: raw.IsCancelled === true ? true : undefined,
    isOnlineMeeting: raw.IsOnlineMeeting === true ? true : undefined,
    categories: Array.isArray(raw.Categories) ? raw.Categories : undefined,
    preview: str(raw.Preview)?.replace(/\s+/g, " ").trim(),
  });
}

export function eventDetail(raw: Raw, options: BodyOptions = {}) {
  return compact({
    ...eventSummary(raw),
    preview: undefined,
    seriesMasterId: raw.SeriesMasterItemId?.Id as string | undefined,
    timeZone: str(raw.StartTimeZoneId),
    sensitivity: raw.Sensitivity && raw.Sensitivity !== "Normal" ? String(raw.Sensitivity).toLowerCase() : undefined,
    importance: importance(raw.Importance),
    reminderMinutesBefore: raw.ReminderIsSet ? raw.ReminderMinutesBeforeStart : undefined,
    isOrganizer: typeof raw.IsOrganizer === "boolean" ? raw.IsOrganizer : undefined,
    responseRequested: raw.IsMeeting && typeof raw.IsResponseRequested === "boolean" ? raw.IsResponseRequested : undefined,
    onlineMeetingUrl: str(raw.OnlineMeetingJoinUrl) ?? str(raw.OnlineMeeting?.JoinUrl),
    attendees: [
      ...(raw.RequiredAttendees ?? []).map((a: Raw) => attendee(a, "required")),
      ...(raw.OptionalAttendees ?? []).map((a: Raw) => attendee(a, "optional")),
      ...(raw.Resources ?? []).map((a: Raw) => attendee(a, "resource")),
    ],
    ...attachmentFields(raw.Attachments),
    ...bodyFields(raw.Body, options),
  });
}

function dictionary(entries: Raw[] | undefined, valueKey: string): Record<string, string> | undefined {
  const out: Record<string, string> = {};
  for (const entry of entries ?? []) {
    const value = str(entry?.[valueKey]);
    if (entry?.Key && value) out[lowerFirst(String(entry.Key))] = value;
  }
  return Object.keys(out).length ? out : undefined;
}

function lowerFirst(text: string): string {
  return text.charAt(0).toLowerCase() + text.slice(1);
}

export function contactSummary(raw: Raw) {
  const emails = (raw.EmailAddresses ?? []).map((e: Raw) => str(e?.EmailAddress)).filter(Boolean);
  return compact({
    id: raw.ItemId?.Id as string,
    displayName: str(raw.DisplayName) ?? str(raw.FileAs),
    emails,
    company: str(raw.CompanyName),
    jobTitle: str(raw.JobTitle),
    mobilePhone: (raw.PhoneNumbers ?? []).find((p: Raw) => p?.Key === "MobilePhone")?.PhoneNumber || undefined,
  });
}

export function contactDetail(raw: Raw) {
  const addresses: Record<string, Raw> = {};
  for (const entry of raw.PhysicalAddresses ?? []) {
    if (!entry?.Key) continue;
    const address = compact({
      street: str(entry.Street),
      city: str(entry.City),
      state: str(entry.State),
      postalCode: str(entry.PostalCode),
      country: str(entry.CountryOrRegion),
    });
    if (Object.keys(address).length) addresses[lowerFirst(String(entry.Key))] = address;
  }
  return compact({
    id: raw.ItemId?.Id as string,
    displayName: str(raw.DisplayName),
    givenName: str(raw.GivenName),
    middleName: str(raw.MiddleName),
    surname: str(raw.Surname),
    nickname: str(raw.Nickname),
    company: str(raw.CompanyName),
    department: str(raw.Department),
    jobTitle: str(raw.JobTitle),
    officeLocation: str(raw.OfficeLocation),
    manager: str(raw.Manager),
    assistantName: str(raw.AssistantName),
    birthday: str(raw.Birthday),
    emails: (raw.EmailAddresses ?? [])
      .filter((e: Raw) => str(e?.EmailAddress))
      .map((e: Raw) => compact({ key: lowerFirst(String(e.Key ?? "")), email: e.EmailAddress, name: str(e.Name) })),
    phones: dictionary(raw.PhoneNumbers, "PhoneNumber"),
    addresses,
    homePage: str(raw.BusinessHomePage) ?? str(raw.PersonalHomePage),
    categories: Array.isArray(raw.Categories) ? raw.Categories : undefined,
    notes: str(raw.Body?.Value)?.trim(),
  });
}

export function persona(raw: Raw) {
  const primary = person(raw.EmailAddress);
  const others = people(raw.EmailAddresses)
    .map((p) => p.email)
    .filter((email): email is string => Boolean(email) && email?.toLowerCase() !== primary?.email?.toLowerCase());
  return compact({
    displayName: str(raw.DisplayName),
    email: primary?.email,
    otherEmails: [...new Set(others)],
    givenName: str(raw.GivenName),
    surname: str(raw.Surname),
    jobTitle: str(raw.Title),
    department: str(raw.Department),
    company: str(raw.CompanyName),
    office: str(raw.OfficeLocation) ?? str(raw.WorkCity),
    phone: str(raw.BusinessPhoneNumber) ?? str(raw.MobilePhone),
    type: str(raw.PersonaTypeString),
  });
}
