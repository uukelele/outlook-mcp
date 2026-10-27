import { z } from "zod";
import { addresses, bodyFormat, defineTool, folderRef, maxBodyChars, type ToolRegistry } from "./helpers.js";

const TIME_HELP =
  "ISO 8601. With an offset (2026-11-02T14:00:00Z or +01:00) it is exact; without one it is local time in timeZone, or in the mailbox time zone if timeZone is omitted";

function optionalShape<S extends Record<string, z.ZodType>>(shape: S): { [K in keyof S]: z.ZodOptional<S[K]> } {
  return Object.fromEntries(Object.entries(shape).map(([key, schema]) => [key, schema.optional()])) as {
    [K in keyof S]: z.ZodOptional<S[K]>;
  };
}

export function registerCalendarTools(registry: ToolRegistry): void {
  const { calendar } = registry.services;

  const recurrence = z
    .object({
      pattern: z.enum(["daily", "weekly", "monthly", "yearly"]).describe("Repeat pattern; monthly/yearly repeat on the start date's day"),
      interval: z.number().int().min(1).max(99).optional().describe("Every N days/weeks/months (default 1)"),
      daysOfWeek: z.array(z.string()).optional().describe('Weekly only: e.g. ["monday", "wednesday"] (default: the start day)'),
      until: z.string().optional().describe("Last date of the series, YYYY-MM-DD"),
      count: z.number().int().min(1).max(999).optional().describe("Number of occurrences (instead of until)"),
    })
    .describe("Make the event repeat");

  const eventFields = {
    subject: z.string().min(1).describe("Title"),
    start: z.string().describe(`Start. ${TIME_HELP}. For all-day events a date (YYYY-MM-DD)`),
    end: z.string().describe(`End. ${TIME_HELP}. For all-day events the day after the last day (exclusive)`),
    timeZone: z.string().optional().describe('IANA time zone for start/end without an offset, e.g. "Europe/London"'),
    isAllDay: z.boolean().optional().describe("All-day event"),
    location: z.string().optional().describe("Location text"),
    body: z.string().optional().describe("Description / agenda"),
    bodyFormat: bodyFormat.optional(),
    attendees: addresses("Required attendees").optional(),
    optionalAttendees: addresses("Optional attendees").optional(),
    reminderMinutesBefore: z.number().int().min(0).max(40_320).nullable().optional().describe("Reminder in minutes before start; null disables the reminder"),
    showAs: z.enum(["free", "tentative", "busy", "oof", "workingElsewhere"]).optional().describe("Free/busy status"),
    sensitivity: z.enum(["normal", "personal", "private", "confidential"]).optional(),
    categories: z.array(z.string()).optional().describe("Category names"),
  };

  defineTool(registry, {
    name: "list_calendars",
    title: "List calendars",
    kind: "read",
    description: "List the mailbox's calendars (the default one is wellKnownName \"calendar\").",
    input: {},
    run: async () => ({ calendars: await calendar.listCalendars() }),
  });

  defineTool(registry, {
    name: "list_events",
    title: "List events",
    kind: "read",
    description:
      "List calendar events (recurring events are expanded into occurrences) that overlap a time range, sorted by start. " +
      "Defaults to the next 7 days from today (UTC). Returned times are UTC.",
    input: {
      start: z.string().optional().describe("Range start (ISO 8601; without an offset it is UTC unless timeZone is given)"),
      end: z.string().optional().describe("Range end (default: start + 7 days)"),
      timeZone: z.string().optional().describe('IANA time zone for start/end without an offset, e.g. "Europe/London"'),
      calendar: folderRef("Calendar (default: the main calendar)").optional(),
      limit: z.number().int().min(1).max(500).optional().describe("Maximum events (default 100)"),
    },
    run: (args) => calendar.listEvents(args),
  });

  defineTool(registry, {
    name: "get_event",
    title: "Get event",
    kind: "read",
    description: "Get full event details: description, attendees with their responses, organizer, online meeting link, recurrence master.",
    input: { id: z.string().min(1).describe("Event ID"), format: bodyFormat.optional(), maxBodyChars: maxBodyChars.optional() },
    run: ({ id, format, maxBodyChars }) => calendar.getEvent(id, { format, maxBodyChars: maxBodyChars ?? 20_000 }),
  });

  defineTool(registry, {
    name: "create_event",
    title: "Create event",
    kind: "send",
    description:
      "Create a calendar event. If attendees are given this is a meeting and invitations are emailed to them immediately. " +
      "Set isOnlineMeeting to add a Microsoft Teams link (if the tenant allows it).",
    input: {
      ...eventFields,
      end: eventFields.end.optional(),
      isOnlineMeeting: z.boolean().optional().describe("Add a Teams meeting link"),
      recurrence: recurrence.optional(),
      calendar: folderRef("Calendar (default: the main calendar)").optional(),
    },
    run: (args) => calendar.createEvent(args),
  });

  defineTool(registry, {
    name: "update_event",
    title: "Update event",
    kind: "send",
    description:
      "Change an event you organize. Only given fields change; attendee lists replace the existing ones. " +
      "For meetings, attendees are sent an update unless notifyAttendees is false.",
    input: {
      id: z.string().min(1).describe("Event ID"),
      ...optionalShape(eventFields),
      notifyAttendees: z.boolean().optional().describe("Send updates to attendees (default true)"),
    },
    run: ({ id, notifyAttendees, ...rest }) => calendar.updateEvent(id, rest, notifyAttendees !== false),
  });

  defineTool(registry, {
    name: "delete_event",
    title: "Delete or cancel event",
    kind: "delete",
    description:
      "Delete an event. For meetings you organize, attendees are sent a cancellation unless notifyAttendees is false. " +
      "To decline someone else's meeting, use respond_to_event instead.",
    input: {
      id: z.string().min(1).describe("Event ID"),
      notifyAttendees: z.boolean().optional().describe("Send cancellations to attendees (default true)"),
      permanent: z.boolean().optional().describe("Skip Deleted Items (default false)"),
    },
    run: ({ id, notifyAttendees, permanent }) => calendar.deleteEvent(id, { notifyAttendees, permanent }),
  });

  defineTool(registry, {
    name: "respond_to_event",
    title: "Respond to invitation",
    kind: "send",
    description: "Accept, tentatively accept or decline a meeting invitation, optionally with a message to the organizer.",
    input: {
      id: z.string().min(1).describe("Event ID (or the meeting request message ID)"),
      response: z.enum(["accept", "tentative", "decline"]),
      message: z.string().optional().describe("Note to the organizer"),
      send: z.boolean().optional().describe("Send the response to the organizer (default true)"),
    },
    run: ({ id, response, message, send }) => calendar.respond(id, response, { message, send }),
  });
}
