import { OutlookError } from "../errors.js";
import {
  attendees,
  bodyContent,
  deleteItemField,
  itemId,
  propertyUri,
  setItemField,
  t,
  targetFolder,
  type BodyFormat,
  type Json,
} from "../owa/builders.js";
import { assertBodySuccess, type OwaClient } from "../owa/client.js";
import { eventDetail, eventSummary } from "../owa/normalize.js";
import { addDays, parseDateOnly, parseDateTimeInput, parseRangeInput, toUtcString, type DateTimeValue } from "../owa/time.js";
import type { FolderService } from "./folders.js";
import type { MailboxService } from "./mailbox.js";

export type ShowAs = "free" | "tentative" | "busy" | "oof" | "workingElsewhere";
const SHOW_AS: Record<ShowAs, string> = { free: "Free", tentative: "Tentative", busy: "Busy", oof: "OOF", workingElsewhere: "WorkingElsewhere" };

export type Sensitivity = "normal" | "personal" | "private" | "confidential";

export interface RecurrenceInput {
  pattern: "daily" | "weekly" | "monthly" | "yearly";
  /** Every N days/weeks/months. Ignored for yearly. */
  interval?: number;
  /** For weekly: e.g. ["monday", "wednesday"]. Defaults to the start date's weekday. */
  daysOfWeek?: string[];
  /** Last date (inclusive, YYYY-MM-DD) of the series. */
  until?: string;
  /** Number of occurrences. */
  count?: number;
}

export interface EventInput {
  subject?: string;
  start?: string;
  end?: string;
  /** IANA zone for start/end values without an offset; otherwise the mailbox time zone is used. */
  timeZone?: string;
  isAllDay?: boolean;
  location?: string;
  body?: string;
  bodyFormat?: BodyFormat;
  attendees?: string[];
  optionalAttendees?: string[];
  reminderMinutesBefore?: number | null;
  showAs?: ShowAs;
  sensitivity?: Sensitivity;
  categories?: string[];
  isOnlineMeeting?: boolean;
  recurrence?: RecurrenceInput;
  calendar?: string;
}

export type EventResponse = "accept" | "tentative" | "decline";
const RESPONSE_ITEM: Record<EventResponse, string> = { accept: "AcceptItem", tentative: "TentativelyAcceptItem", decline: "DeclineItem" };

const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

interface ResolvedTimes {
  start?: string;
  end?: string;
  /** Time zone for the request's TimeZoneContext. */
  contextZone: string;
  /** Windows zone ID to store on the event. */
  eventZone: string;
  startDate?: string;
}

export class CalendarService {
  constructor(
    private readonly client: OwaClient,
    private readonly folders: FolderService,
    private readonly mailbox: MailboxService,
  ) {}

  async listCalendars() {
    return (await this.folders.list("calendar")).map(({ id, name, path, wellKnownName, totalCount }) => ({
      id,
      name,
      path,
      wellKnownName,
      itemCount: totalCount,
    }));
  }

  async listEvents(options: { start?: string; end?: string; calendar?: string; limit?: number; timeZone?: string } = {}) {
    const start = options.start ? parseRangeInput(options.start, "start", options.timeZone) : startOfToday();
    const end = options.end ? parseRangeInput(options.end, "end", options.timeZone) : new Date(start.getTime() + 7 * 86_400_000);
    if (end <= start) throw new OutlookError("end must be after start");
    const calendarId = await this.folders.resolve(options.calendar ?? "calendar");
    const result = await this.client.call("GetCalendarView", {
      CalendarId: targetFolder(calendarId),
      RangeStart: toUtcString(start),
      RangeEnd: toUtcString(end),
    });
    const body = assertBodySuccess("GetCalendarView", result);
    const events = ((body.Items as Json[] | undefined) ?? [])
      .map(eventSummary)
      .sort((a, b) => (a.start ?? "").localeCompare(b.start ?? ""));
    const limit = Math.max(1, Math.min(options.limit ?? 100, 500));
    return {
      range: { start: toUtcString(start), end: toUtcString(end) },
      total: events.length,
      truncated: events.length > limit ? true : undefined,
      events: events.slice(0, limit),
    };
  }

  async getEvent(id: string, options: { format?: BodyFormat; maxBodyChars?: number } = {}) {
    const [message] = await this.client.callForMessages("GetCalendarEvent", {
      EventIds: [itemId(id)],
      ItemShape: {
        __type: t("ItemResponseShape"),
        BaseShape: "IdOnly",
        BodyType: options.format === "html" ? "HTML" : "Text",
        ...(options.format === "html" ? { FilterHtmlContent: true } : {}),
      },
    });
    const raw = (message?.Items as Json[] | undefined)?.[0];
    if (!raw) throw new OutlookError("Event not found");
    return eventDetail(raw, { maxChars: options.maxBodyChars });
  }

  async createEvent(input: EventInput) {
    if (!input.subject) throw new OutlookError("subject is required");
    if (!input.start) throw new OutlookError("start is required");
    const times = await this.resolveTimes(input, true);
    const item: Json = { __type: t("CalendarItem"), Subject: input.subject };
    if (input.body !== undefined) item.Body = bodyContent(input.body, input.bodyFormat);
    item.Start = times.start;
    item.End = times.end;
    item.StartTimeZoneId = times.eventZone;
    item.EndTimeZoneId = times.eventZone;
    item.IsAllDayEvent = Boolean(input.isAllDay);
    item.FreeBusyType = SHOW_AS[input.showAs ?? (input.isAllDay ? "free" : "busy")];
    if (input.sensitivity) item.Sensitivity = capitalize(input.sensitivity);
    if (input.reminderMinutesBefore === null) item.ReminderIsSet = false;
    else if (input.reminderMinutesBefore !== undefined) {
      item.ReminderIsSet = true;
      item.ReminderMinutesBeforeStart = input.reminderMinutesBefore;
    }
    if (input.location) item.Location = enhancedLocation(input.location);
    if (input.attendees?.length) item.RequiredAttendees = attendees(input.attendees);
    if (input.optionalAttendees?.length) item.OptionalAttendees = attendees(input.optionalAttendees);
    if (input.categories?.length) item.Categories = input.categories;
    if (input.isOnlineMeeting) {
      item.IsOnlineMeeting = true;
      item.OnlineMeetingProvider = "TeamsForBusiness";
    }
    if (input.recurrence) item.Recurrence = recurrence(input.recurrence, times.startDate!);

    const calendarId = await this.folders.resolve(input.calendar ?? "calendar");
    const invites = Boolean(input.attendees?.length || input.optionalAttendees?.length);
    const [message] = await this.client.callForMessages(
      "CreateCalendarEvent",
      {
        SavedItemFolderId: targetFolder(calendarId),
        SendMeetingInvitations: invites ? "SendToAllAndSaveCopy" : "SendToNone",
        Items: [item],
      },
      { jsonType: "CreateItem", timeZone: times.contextZone },
    );
    const id = (((message?.Items as Json[] | undefined)?.[0]?.ItemId as Json | undefined)?.Id as string | undefined) ?? undefined;
    if (!id) throw new OutlookError("CreateCalendarEvent succeeded but returned no event ID");
    return { invitationsSent: invites, event: await this.getEvent(id, { maxBodyChars: 2000 }) };
  }

  async updateEvent(id: string, input: EventInput, notifyAttendees = true) {
    const updates: Json[] = [];
    const set = (uri: string, values: Json) => updates.push(setItemField(propertyUri(uri), "CalendarItem", values));
    const times = input.start || input.end ? await this.resolveTimes(input, false) : undefined;

    if (input.subject !== undefined) set("item:Subject", { Subject: input.subject });
    if (input.body !== undefined) set("item:Body", { Body: bodyContent(input.body, input.bodyFormat) });
    if (times?.start) {
      set("calendar:Start", { Start: times.start });
      set("calendar:StartTimeZoneId", { StartTimeZoneId: times.eventZone });
    }
    if (times?.end) {
      set("calendar:End", { End: times.end });
      set("calendar:EndTimeZoneId", { EndTimeZoneId: times.eventZone });
    }
    if (input.isAllDay !== undefined) set("calendar:IsAllDayEvent", { IsAllDayEvent: input.isAllDay });
    if (input.location !== undefined) {
      if (input.location) set("EnhancedLocation", { Location: enhancedLocation(input.location) });
      else updates.push(deleteItemField(propertyUri("EnhancedLocation")));
    }
    if (input.reminderMinutesBefore === null) set("item:ReminderIsSet", { ReminderIsSet: false });
    else if (input.reminderMinutesBefore !== undefined) {
      set("item:ReminderIsSet", { ReminderIsSet: true });
      set("item:ReminderMinutesBeforeStart", { ReminderMinutesBeforeStart: input.reminderMinutesBefore });
    }
    if (input.sensitivity) set("item:Sensitivity", { Sensitivity: capitalize(input.sensitivity) });
    if (input.categories !== undefined) {
      if (input.categories.length) set("item:Categories", { Categories: input.categories });
      else updates.push(deleteItemField(propertyUri("item:Categories")));
    }
    if (input.attendees !== undefined) {
      if (input.attendees.length) set("calendar:RequiredAttendees", { RequiredAttendees: attendees(input.attendees) });
      else updates.push(deleteItemField(propertyUri("calendar:RequiredAttendees")));
    }
    if (input.optionalAttendees !== undefined) {
      if (input.optionalAttendees.length) set("calendar:OptionalAttendees", { OptionalAttendees: attendees(input.optionalAttendees) });
      else updates.push(deleteItemField(propertyUri("calendar:OptionalAttendees")));
    }
    if (!updates.length) throw new OutlookError("Nothing to update");

    await this.client.callForMessages(
      "UpdateCalendarEvent",
      {
        EventId: itemId(id),
        ItemChange: { __type: t("ItemChange"), ItemId: itemId(id), Updates: updates },
        SendCalendarInvitationsOrCancellations: notifyAttendees ? "SendToAllAndSaveCopy" : "SendToNone",
      },
      { timeZone: times?.contextZone },
    );
    return this.getEvent(id, { maxBodyChars: 2000 });
  }

  async deleteEvent(id: string, options: { notifyAttendees?: boolean; permanent?: boolean } = {}) {
    await this.client.callForMessages("DeleteItem", {
      ItemIds: [itemId(id)],
      DeleteType: options.permanent ? "HardDelete" : "MoveToDeletedItems",
      SendMeetingCancellations: options.notifyAttendees === false ? "SendToNone" : "SendToAllAndSaveCopy",
    });
    return { deleted: true };
  }

  async respond(id: string, response: EventResponse, options: { message?: string; send?: boolean } = {}) {
    const item: Json = { __type: t(RESPONSE_ITEM[response]), ReferenceItemId: itemId(id) };
    if (options.message) item.Body = bodyContent(options.message, "text");
    const send = options.send !== false;
    await this.client.callForMessages("CreateItem", {
      MessageDisposition: send ? "SendAndSaveCopy" : "SaveOnly",
      Items: [item],
    });
    return { response, responseSent: send };
  }

  private async resolveTimes(input: EventInput, creating: boolean): Promise<ResolvedTimes> {
    const mailboxZone = await this.mailbox.timeZone();
    if (input.isAllDay) {
      if (input.timeZone) throw new OutlookError("timeZone does not apply to all-day events; pass dates like 2026-11-02");
      const startDate = input.start ? parseDateOnly(input.start, "start") : undefined;
      let endDate = input.end ? parseDateOnly(input.end, "end") : undefined;
      if (creating && !endDate) endDate = addDays(startDate!, 1);
      if (startDate && endDate && endDate <= startDate) {
        throw new OutlookError("For all-day events, end is exclusive and must be after start (e.g. start 2026-11-02, end 2026-11-03 for one day)");
      }
      return {
        start: startDate && `${startDate}T00:00:00`,
        end: endDate && `${endDate}T00:00:00`,
        contextZone: mailboxZone,
        eventZone: mailboxZone,
        startDate,
      };
    }
    const start: DateTimeValue | undefined = input.start ? parseDateTimeInput(input.start, "start", input.timeZone) : undefined;
    let end: DateTimeValue | undefined = input.end ? parseDateTimeInput(input.end, "end", input.timeZone) : undefined;
    if (creating && !end) {
      if (start!.floating) throw new OutlookError("end is required (or give start with an offset/timeZone to default to a 30-minute event)");
      end = { value: toUtcString(new Date(Date.parse(start!.value) + 30 * 60_000)), floating: false };
    }
    if (start && end && start.floating !== end.floating) {
      throw new OutlookError("start and end must both include an offset (e.g. Z or +01:00) or both omit it");
    }
    const floating = Boolean(start?.floating ?? end?.floating);
    if (start && end) {
      const a = floating ? `${start.value}Z` : start.value;
      const b = floating ? `${end.value}Z` : end.value;
      if (Date.parse(b) <= Date.parse(a)) throw new OutlookError("end must be after start");
    }
    return {
      start: start?.value,
      end: end?.value,
      contextZone: floating ? mailboxZone : "UTC",
      eventZone: mailboxZone,
      startDate: input.start?.trim().slice(0, 10),
    };
  }
}

function enhancedLocation(name: string): Json {
  return {
    __type: t("EnhancedLocation"),
    DisplayName: name,
    PostalAddress: { __type: t("PersonaPostalAddress"), Type: "Business", LocationSource: "None" },
  };
}

function recurrence(input: RecurrenceInput, startDate: string): Json {
  const start = new Date(`${startDate}T00:00:00Z`);
  if (Number.isNaN(start.getTime())) throw new OutlookError("Recurring events need a start that begins with a YYYY-MM-DD date");
  const interval = Math.max(1, Math.floor(input.interval ?? 1));
  let pattern: Json;
  switch (input.pattern) {
    case "daily":
      pattern = { __type: t("DailyRecurrence"), Interval: interval };
      break;
    case "weekly": {
      const days = (input.daysOfWeek?.length ? input.daysOfWeek : [WEEKDAYS[start.getUTCDay()]!]).map(weekday);
      pattern = { __type: t("WeeklyRecurrence"), Interval: interval, DaysOfWeek: days.join(" "), FirstDayOfWeek: "Monday" };
      break;
    }
    case "monthly":
      pattern = { __type: t("AbsoluteMonthlyRecurrence"), Interval: interval, DayOfMonth: start.getUTCDate() };
      break;
    case "yearly":
      pattern = { __type: t("AbsoluteYearlyRecurrence"), DayOfMonth: start.getUTCDate(), Month: MONTHS[start.getUTCMonth()] };
      break;
  }
  if (input.until && input.count) throw new OutlookError("Give recurrence.until or recurrence.count, not both");
  let range: Json;
  if (input.until) {
    const until = parseDateOnly(input.until, "recurrence.until");
    if (until < startDate) throw new OutlookError("recurrence.until is before the start date");
    range = { __type: t("EndDateRecurrence"), StartDate: startDate, EndDate: until };
  } else if (input.count) {
    range = { __type: t("NumberedRecurrence"), StartDate: startDate, NumberOfOccurrences: Math.max(1, Math.floor(input.count)) };
  } else {
    range = { __type: t("NoEndRecurrence"), StartDate: startDate };
  }
  return { __type: t("RecurrenceType"), RecurrencePattern: pattern, RecurrenceRange: range };
}

function weekday(input: string): string {
  const match = WEEKDAYS.find((day) => day.toLowerCase().startsWith(input.trim().toLowerCase().slice(0, 3)));
  if (!match || input.trim().length < 2) throw new OutlookError(`"${input}" is not a weekday`);
  return match;
}

function startOfToday(): Date {
  const now = new Date();
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
}

function capitalize(text: string): string {
  return text.charAt(0).toUpperCase() + text.slice(1);
}
