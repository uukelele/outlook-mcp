import { describe, expect, it } from "vitest";
import { compact, contactDetail, eventDetail, eventSummary, messageDetail, messageSummary, person, persona } from "../src/owa/normalize.js";

describe("person", () => {
  it("prefers the SMTP address over Exchange legacy DNs", () => {
    expect(
      person({
        Mailbox: { Name: "Pat", EmailAddress: "/O=EXCHANGELABS/OU=EXCHANGE/CN=RECIPIENTS/CN=PAT", RoutingType: "EX", SmtpEmailAddress: "pat@contoso.com" },
      }),
    ).toEqual({ name: "Pat", email: "pat@contoso.com" });
    expect(person({ Name: "x@contoso.com", EmailAddress: "x@contoso.com", RoutingType: "SMTP" })).toEqual({ name: undefined, email: "x@contoso.com" });
  });
});

describe("compact", () => {
  it("drops empty values recursively", () => {
    expect(compact({ a: 1, b: null, c: "", d: [], e: { f: undefined }, g: [null, 2], h: false })).toEqual({ a: 1, g: [2], h: false });
  });
});

describe("messages", () => {
  const raw = {
    __type: "Message:#Exchange",
    ItemId: { Id: "M1", ChangeKey: "CK" },
    ConversationId: { Id: "C1" },
    Subject: "Hello",
    From: { Mailbox: { Name: "Ann", EmailAddress: "ann@contoso.com", RoutingType: "SMTP" } },
    ToRecipients: [{ Name: "Me", EmailAddress: "me@contoso.com", RoutingType: "SMTP" }],
    DisplayTo: "Me",
    DateTimeReceived: "2026-10-23T09:00:00Z",
    IsRead: false,
    HasAttachments: true,
    Importance: "High",
    Flag: { FlagStatus: "Flagged" },
    Categories: ["Blue"],
    Preview: "Hi\r\n\r\nthere",
    Body: { BodyType: "Text", Value: "Hi\r\n\r\n\r\n\r\nthere" },
    Attachments: [
      { __type: "FileAttachment:#Exchange", AttachmentId: { Id: "A1" }, Name: "a.pdf", ContentType: "application/pdf", Size: 10 },
      { __type: "ItemAttachment:#Exchange", AttachmentId: { Id: "A2" }, Name: "fwd" },
      { __type: "FileAttachment:#Exchange", AttachmentId: { Id: "A3" }, Name: "image001.png", IsInline: true },
    ],
  };

  it("summarizes", () => {
    expect(messageSummary(raw)).toEqual({
      id: "M1",
      conversationId: "C1",
      subject: "Hello",
      from: "Ann <ann@contoso.com>",
      to: "Me",
      receivedAt: "2026-10-23T09:00:00Z",
      isRead: false,
      hasAttachments: true,
      importance: "high",
      flag: "flagged",
      categories: ["Blue"],
      preview: "Hi there",
    });
  });

  it("details with body and attachments, truncating long bodies", () => {
    const detail = messageDetail(raw, { maxChars: 3 });
    expect(detail.to).toEqual([{ name: "Me", email: "me@contoso.com" }]);
    expect(detail.attachments).toEqual([
      { id: "A1", name: "a.pdf", contentType: "application/pdf", size: 10 },
      { id: "A2", name: "fwd", kind: "item" },
    ]);
    expect(detail.inlineImages).toBe(1);
    expect(detail).toMatchObject({ bodyFormat: "text", bodyTruncated: true });
    expect(typeof detail.body).toBe("string");
    expect(detail.body!.startsWith("Hi\n")).toBe(true);
  });

  it("prefers UniqueBody for thread items", () => {
    expect(messageDetail({ ...raw, UniqueBody: { BodyType: "Text", Value: "new part" } }).body).toBe("new part");
  });
});

describe("events", () => {
  const raw = {
    ItemId: { Id: "E1" },
    Subject: "Sync",
    Start: "2026-10-24T10:00:00Z",
    End: "2026-10-24T10:30:00Z",
    Location: { DisplayName: "Room 1" },
    Organizer: { Mailbox: { Name: "Ann", EmailAddress: "ann@contoso.com", RoutingType: "SMTP" } },
    IsMeeting: true,
    ResponseType: "Accept",
    FreeBusyType: "Busy",
    RequiredAttendees: [{ Mailbox: { Name: "Bob", EmailAddress: "bob@contoso.com" }, ResponseType: "Decline" }],
    OptionalAttendees: [{ Mailbox: { EmailAddress: "cy@contoso.com" }, ResponseType: "Unknown" }],
    ReminderIsSet: true,
    ReminderMinutesBeforeStart: 15,
    StartTimeZoneId: "GMT Standard Time",
    Body: { BodyType: "Text", Value: "Agenda" },
  };

  it("summarizes", () => {
    expect(eventSummary(raw)).toEqual({
      id: "E1",
      subject: "Sync",
      start: "2026-10-24T10:00:00Z",
      end: "2026-10-24T10:30:00Z",
      location: "Room 1",
      organizer: "Ann <ann@contoso.com>",
      myResponse: "accepted",
      showAs: "Busy",
      isMeeting: true,
    });
  });

  it("details attendees and responses", () => {
    const detail = eventDetail(raw);
    expect(detail.attendees).toEqual([
      { name: "Bob", email: "bob@contoso.com", type: "required", response: "declined" },
      { email: "cy@contoso.com", type: "optional", response: "unknown" },
    ]);
    expect(detail).toMatchObject({ reminderMinutesBefore: 15, timeZone: "GMT Standard Time", body: "Agenda", bodyFormat: "text" });
  });
});

describe("contacts and people", () => {
  it("normalizes indexed contact fields", () => {
    const detail = contactDetail({
      ItemId: { Id: "P1" },
      DisplayName: "Probe Person",
      EmailAddresses: [{ Key: "EmailAddress1", Name: "Probe", EmailAddress: "probe@example.com", RoutingType: "SMTP" }],
      PhoneNumbers: [
        { Key: "MobilePhone", PhoneNumber: "+44 7000 000000" },
        { Key: "CarPhone", PhoneNumber: "" },
      ],
      PhysicalAddresses: [{ Key: "Business", Street: "1 High St", City: "Town", CountryOrRegion: "UK" }],
    });
    expect(detail).toEqual({
      id: "P1",
      displayName: "Probe Person",
      emails: [{ key: "emailAddress1", email: "probe@example.com", name: "Probe" }],
      phones: { mobilePhone: "+44 7000 000000" },
      addresses: { business: { street: "1 High St", city: "Town", country: "UK" } },
    });
  });

  it("normalizes directory personas", () => {
    expect(
      persona({
        DisplayName: "Doe, Jane",
        EmailAddress: { Name: "Doe, Jane", EmailAddress: "jane@contoso.com", RoutingType: "SMTP" },
        EmailAddresses: [{ EmailAddress: "jane@contoso.com" }, { EmailAddress: "j.doe@contoso.com" }],
        GivenName: "Jane",
        Surname: "Doe",
        WorkCity: "London",
        PersonaTypeString: "Person",
      }),
    ).toEqual({
      displayName: "Doe, Jane",
      email: "jane@contoso.com",
      otherEmails: ["j.doe@contoso.com"],
      givenName: "Jane",
      surname: "Doe",
      office: "London",
      type: "Person",
    });
  });
});
