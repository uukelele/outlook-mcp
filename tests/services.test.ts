import { describe, expect, it } from "vitest";
import { typeKeysFirst } from "../src/owa/builders.js";
import { createServices } from "../src/server.js";
import { FOLDERS, ID, failure, fakeOwa, findFolderResponse, messages, success, testAuth, testConfig } from "./fake-owa.js";

const USER_CONFIG = {
  SessionSettings: { UserEmailAddress: "me@contoso.com", UserDisplayName: "Me", UserPrincipalName: "me@contoso.com" },
  UserOptions: { TimeZone: "GMT Standard Time" },
  MasterCategoryList: { MasterList: [{ Name: "Blue", Color: 7, Id: "c1" }, { Name: "Admin", Color: -1 }] },
};

function setup(handlers: Record<string, (body: any, call: any) => unknown> = {}) {
  const owa = fakeOwa({
    FindFolder: () => findFolderResponse(),
    GetOwaUserConfiguration: () => USER_CONFIG,
    ...handlers,
  });
  const services = createServices(testConfig(), { auth: testAuth, fetch: owa.fetch });
  return { owa, ...services };
}

const created = (id: string) => messages(success({ Items: [{ ItemId: { Id: id, ChangeKey: "CK" } }] }));
const msg = (n: number, extra: Record<string, unknown> = {}) => ({
  ItemId: { Id: ID(n) },
  Subject: `Message ${n}`,
  From: { Mailbox: { Name: `Sender ${n}`, EmailAddress: `s${n}@contoso.com` } },
  DateTimeReceived: `2026-10-${String(30 - n).padStart(2, "0")}T09:00:00Z`,
  IsRead: n % 2 === 0,
  Preview: `preview ${n}`,
  ...extra,
});

describe("FolderService", () => {
  it("resolves well-known names, paths, display names and IDs", async () => {
    const { folders } = setup();
    expect(await folders.resolve("sent")).toBe("sentitems");
    expect(await folders.resolve("Inbox/Projects")).toBe(FOLDERS.projects.Id);
    expect(await folders.resolve("projects")).toBe(FOLDERS.projects.Id);
    expect(await folders.resolve(FOLDERS.archive.Id)).toBe(FOLDERS.archive.Id);
    await expect(folders.resolve("Nope")).rejects.toThrow(/not found/);
  });

  it("lists mail folders with paths and hides system folders", async () => {
    const { folders } = setup();
    const list = await folders.list("mail");
    expect(list.map((f) => f.path)).toEqual(["Archive", "Inbox", "Inbox/Projects", "Sent Items"]);
    expect((await folders.list("calendar")).map((f) => f.name)).toEqual(["Calendar"]);
  });
});

describe("MailService reading", () => {
  it("lists with server-side filter, sort and paging", async () => {
    const { mail, owa } = setup({
      FindItem: () => messages(success({ RootFolder: { Items: [msg(1), msg(2)], TotalItemsInView: 40, IncludesLastItemInRange: false } })),
    });
    const result = await mail.list({ folder: "Inbox/Projects", filter: "unread", limit: 2, offset: 10 });
    const body = owa.of("FindItem")[0]!.body.Body;
    expect(body.ParentFolderIds).toEqual([{ __type: "FolderId:#Exchange", Id: FOLDERS.projects.Id }]);
    expect(body.ViewFilter).toBe("Unread");
    expect(body.Paging).toEqual({ __type: "IndexedPageView:#Exchange", BasePoint: "Beginning", Offset: 10, MaxEntriesReturned: 2 });
    expect(body.SortOrder[0]).toMatchObject({ Order: "Descending", Path: { FieldURI: "item:DateTimeReceived" } });
    expect(result).toMatchObject({ total: 40, count: 2, nextOffset: 12 });
    expect(result.messages[0]).toMatchObject({ id: ID(1), from: "Sender 1 <s1@contoso.com>" });
  });

  it("searches across folders client-side and stops at the 'after' date", async () => {
    const pages: Record<string, unknown[]> = {
      [FOLDERS.inbox.Id]: [msg(1, { Subject: "Budget review" }), msg(2), msg(5, { Subject: "Old budget" })],
      [FOLDERS.projects.Id]: [msg(3, { Subject: "Project budget", HasAttachments: true })],
      [FOLDERS.sent.Id]: [msg(4, { Preview: "re: the budget numbers" })],
    };
    const { mail, owa } = setup({
      FindItem: (body) => {
        const folder = body.Body.ParentFolderIds[0].Id;
        return messages(success({ RootFolder: { Items: pages[folder] ?? [], IncludesLastItemInRange: true } }));
      },
    });
    const result = await mail.search({ text: "budget", after: "2026-10-26" });
    expect(result.foldersSearched).toEqual(["Inbox", "Inbox/Projects", "Sent Items"]);
    expect(result.messages.map((m) => m.subject)).toEqual(["Budget review", "Project budget", "Message 4"]);
    expect(owa.of("FindItem").length).toBe(3);

    const withAttachments = await mail.search({ text: "budget", hasAttachments: true, folders: ["all"] });
    expect(withAttachments.messages.map((m) => m.folder)).toEqual(["Inbox/Projects"]);
  });

  it("gets a message as text and can mark it read", async () => {
    const { mail, owa } = setup({
      GetItem: () => messages(success({ Items: [msg(1, { IsRead: false, Body: { BodyType: "Text", Value: "Hello" } })] })),
      UpdateItem: () => messages(success()),
    });
    const detail = await mail.get(ID(1), { markAsRead: true });
    expect(owa.of("GetItem")[0]!.body.Body.ItemShape).toMatchObject({ BaseShape: "IdOnly", BodyType: "Text" });
    expect(detail).toMatchObject({ isRead: true, body: "Hello", bodyFormat: "text" });
    expect(owa.of("UpdateItem")[0]!.body.Body.ItemChanges[0].Updates[0].Path.FieldURI).toBe("IsRead");
  });

  it("assembles a conversation oldest-first", async () => {
    const { mail, owa } = setup({
      GetItem: () => messages(success({ Items: [{ ConversationId: { Id: "CONV" } }] })),
      GetConversationItems: () =>
        messages(success({ Conversation: { ConversationNodes: [{ Items: [msg(1)] }, { Items: [msg(3)] }] } })),
    });
    const thread = await mail.thread({ messageId: ID(1) });
    expect(owa.of("GetConversationItems")[0]!.body.Body.Conversations[0].ConversationId.Id).toBe("CONV");
    expect(thread.messages.map((m) => m.id)).toEqual([ID(3), ID(1)]);
  });
});

describe("MailService writing (payload shapes verified against the live service)", () => {
  it("sends new mail with CreateItem SendAndSaveCopy and flat recipients", async () => {
    const { mail, owa } = setup({ CreateItem: () => messages(success({ Items: [null] })) });
    await mail.send({ to: ["Jane <jane@contoso.com>"], cc: ["bob@contoso.com"], subject: "Hi", body: "<b>x</b>", bodyFormat: "html", importance: "high" });
    const body = owa.of("CreateItem")[0]!.body;
    expect(typeKeysFirst(body)).toBe(true);
    expect(body.Body).toEqual({
      __type: "CreateItemRequest:#Exchange",
      MessageDisposition: "SendAndSaveCopy",
      SavedItemFolderId: { __type: "TargetFolderId:#Exchange", BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "sentitems" } },
      Items: [
        {
          __type: "Message:#Exchange",
          Subject: "Hi",
          Body: { __type: "BodyContentType:#Exchange", BodyType: "HTML", Value: "<b>x</b>" },
          ToRecipients: [{ Name: "Jane", EmailAddress: "jane@contoso.com", RoutingType: "SMTP" }],
          CcRecipients: [{ EmailAddress: "bob@contoso.com", RoutingType: "SMTP" }],
          Importance: "High",
        },
      ],
    });
    await expect(mail.send({ subject: "x", body: "y" })).rejects.toThrow(/recipient/);
  });

  it("sends with attachments via draft → CreateAttachment → UpdateItem(SendAndSaveCopy)", async () => {
    const { mail, owa } = setup({
      CreateItem: () => created("DRAFT"),
      CreateAttachment: () => messages(success({ Attachments: [{ AttachmentId: { Id: "ATT" } }] })),
      GetItem: () => messages(success({ Items: [{ Subject: "Report", IsDraft: true }] })),
      UpdateItem: () => messages(success()),
    });
    await mail.send({ to: ["a@contoso.com"], subject: "Report", body: "see attached", attachments: [{ name: "notes.txt", text: "hello" }] });
    expect(owa.calls.map((c) => c.action)).toEqual(["CreateItem", "CreateAttachment", "GetItem", "UpdateItem"]);
    expect(owa.of("CreateItem")[0]!.body.Body).toMatchObject({ MessageDisposition: "SaveOnly", SavedItemFolderId: { BaseFolderId: { Id: "drafts" } } });
    expect(owa.of("CreateAttachment")[0]!.body.Body).toMatchObject({
      ParentItemId: { Id: "DRAFT" },
      Attachments: [{ __type: "FileAttachment:#Exchange", Name: "notes.txt", ContentType: "text/plain", Content: Buffer.from("hello").toString("base64") }],
    });
    expect(owa.of("UpdateItem")[0]!.body.Body).toMatchObject({
      MessageDisposition: "SendAndSaveCopy",
      ItemChanges: [{ ItemId: { Id: "DRAFT" }, Updates: [{ Path: { FieldURI: "item:Subject" }, Item: { Subject: "Report" } }] }],
    });
  });

  it("refuses to 'send' an item that is not a draft", async () => {
    const { mail } = setup({ GetItem: () => messages(success({ Items: [{ Subject: "x", IsDraft: false }] })) });
    await expect(mail.sendDraft(ID(1))).rejects.toThrow(/not a draft/);
  });

  it("replies with ReplyAllToItem + ShouldIgnoreChangeKey, or saves a draft", async () => {
    const { mail, owa } = setup({ CreateItem: () => created("REPLY") });
    const result = await mail.reply(ID(1), { body: "Thanks!", replyAll: true, send: false });
    expect(result).toEqual({ sent: false, draftId: "REPLY" });
    expect(owa.of("CreateItem")[0]!.body.Body).toEqual({
      __type: "CreateItemRequest:#Exchange",
      MessageDisposition: "SaveOnly",
      SavedItemFolderId: { __type: "TargetFolderId:#Exchange", BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "drafts" } },
      Items: [
        {
          __type: "ReplyAllToItem:#Exchange",
          ReferenceItemId: { __type: "ItemId:#Exchange", Id: ID(1) },
          ShouldIgnoreChangeKey: true,
          NewBodyContent: { __type: "BodyContentType:#Exchange", BodyType: "Text", Value: "Thanks!" },
        },
      ],
    });
  });

  it("forwards with ForwardItem and flat recipients", async () => {
    const { mail, owa } = setup({ CreateItem: () => messages(success({ Items: [null] })) });
    expect(await mail.forward(ID(1), { to: ["x@contoso.com"], body: "FYI", send: true })).toEqual({ sent: true });
    expect(owa.of("CreateItem")[0]!.body.Body.Items[0]).toMatchObject({
      __type: "ForwardItem:#Exchange",
      ToRecipients: [{ EmailAddress: "x@contoso.com", RoutingType: "SMTP" }],
    });
    expect(owa.of("CreateItem")[0]!.body.Body.MessageDisposition).toBe("SendAndSaveCopy");
  });

  it("batches triage updates and reports per-item failures", async () => {
    const { mail, owa } = setup({ UpdateItem: () => messages(success(), failure("ErrorItemNotFound")) });
    const result = await mail.setFlag([ID(1), ID(2)], "flagged");
    expect(result.succeeded).toEqual([ID(1)]);
    expect(result.failed).toEqual([{ id: ID(2), error: expect.stringMatching(/ErrorItemNotFound/) }]);
    expect(owa.of("UpdateItem")[0]!.body.Body.ItemChanges[1].Updates[0]).toEqual({
      __type: "SetItemField:#Exchange",
      Path: { __type: "PropertyUri:#Exchange", FieldURI: "item:Flag" },
      Item: { __type: "Message:#Exchange", Flag: { __type: "FlagType:#Exchange", FlagStatus: "Flagged" } },
    });
  });

  it("adds categories without dropping existing ones", async () => {
    const { mail, owa } = setup({
      GetItem: () => messages(success({ Items: [{ Categories: ["Blue"] }] })),
      UpdateItem: () => messages(success()),
    });
    await mail.setCategories([ID(1)], ["Admin", "blue"], "add");
    expect(owa.of("UpdateItem")[0]!.body.Body.ItemChanges[0].Updates[0].Item.Categories).toEqual(["Blue", "Admin", "blue"]);
  });

  it("moves to a resolved folder and deletes to Deleted Items", async () => {
    const { mail, owa } = setup({ MoveItem: () => messages(success()), DeleteItem: () => messages(success()) });
    await mail.move([ID(1)], "Archive");
    expect(owa.of("MoveItem")[0]!.body.Body.ToFolderId.BaseFolderId).toEqual({ __type: "DistinguishedFolderId:#Exchange", Id: "archive" });
    await mail.delete([ID(1)], false);
    expect(owa.of("DeleteItem")[0]!.body.Body).toMatchObject({ DeleteType: "MoveToDeletedItems", ItemIds: [{ Id: ID(1) }] });
  });
});

describe("CalendarService", () => {
  const event = { ItemId: { Id: "EV" }, Subject: "Planning", Start: "2026-11-02T10:00:00Z", End: "2026-11-02T10:30:00Z" };

  it("lists a calendar view in UTC, sorted by start", async () => {
    const { calendar, owa } = setup({
      GetCalendarView: () => ({ Body: { ResponseClass: "Success", Items: [{ ...event, ItemId: { Id: "B" }, Start: "2026-11-03T09:00:00Z" }, event] } }),
    });
    const result = await calendar.listEvents({ start: "2026-11-01", end: "2026-11-08" });
    expect(owa.of("GetCalendarView")[0]!.body.Body).toEqual({
      __type: "GetCalendarViewRequest:#Exchange",
      CalendarId: { __type: "TargetFolderId:#Exchange", BaseFolderId: { __type: "DistinguishedFolderId:#Exchange", Id: "calendar" } },
      RangeStart: "2026-11-01T00:00:00Z",
      RangeEnd: "2026-11-08T00:00:00Z",
    });
    expect(result.events.map((e) => e.id)).toEqual(["EV", "B"]);
  });

  it("creates meetings with CreateCalendarEvent, wall-clock times in the mailbox zone and wrapped attendees", async () => {
    const { calendar, owa } = setup({
      CreateCalendarEvent: () => created("EV"),
      GetCalendarEvent: () => messages(success({ Items: [event] })),
    });
    const result = await calendar.createEvent({
      subject: "Planning",
      start: "2026-11-02T10:00",
      end: "2026-11-02T10:30",
      location: "Room 1",
      attendees: ["bob@contoso.com"],
      reminderMinutesBefore: 10,
    });
    const call = owa.of("CreateCalendarEvent")[0]!;
    expect(call.body.__type).toBe("CreateItemJsonRequest:#Exchange");
    expect(call.body.Header.TimeZoneContext.TimeZoneDefinition.Id).toBe("GMT Standard Time");
    expect(call.body.Body).toMatchObject({
      __type: "CreateItemRequest:#Exchange",
      SendMeetingInvitations: "SendToAllAndSaveCopy",
      Items: [
        {
          __type: "CalendarItem:#Exchange",
          Start: "2026-11-02T10:00:00",
          End: "2026-11-02T10:30:00",
          StartTimeZoneId: "GMT Standard Time",
          IsAllDayEvent: false,
          FreeBusyType: "Busy",
          ReminderIsSet: true,
          ReminderMinutesBeforeStart: 10,
          Location: { __type: "EnhancedLocation:#Exchange", DisplayName: "Room 1" },
          RequiredAttendees: [{ __type: "AttendeeType:#Exchange", Mailbox: { EmailAddress: "bob@contoso.com", RoutingType: "SMTP" } }],
        },
      ],
    });
    expect(typeKeysFirst(call.body)).toBe(true);
    expect(result).toMatchObject({ invitationsSent: true, event: { id: "EV", subject: "Planning" } });
  });

  it("uses a UTC context for exact times, defaults to 30 minutes, and builds recurrences", async () => {
    const { calendar, owa } = setup({
      CreateCalendarEvent: () => created("EV"),
      GetCalendarEvent: () => messages(success({ Items: [event] })),
    });
    await calendar.createEvent({
      subject: "Standup",
      start: "2026-11-02T09:00:00Z",
      recurrence: { pattern: "weekly", daysOfWeek: ["mon", "wednesday"], count: 6 },
    });
    const body = owa.of("CreateCalendarEvent")[0]!.body;
    expect(body.Header.TimeZoneContext.TimeZoneDefinition.Id).toBe("UTC");
    expect(body.Body.SendMeetingInvitations).toBe("SendToNone");
    expect(body.Body.Items[0]).toMatchObject({
      Start: "2026-11-02T09:00:00Z",
      End: "2026-11-02T09:30:00Z",
      Recurrence: {
        __type: "RecurrenceType:#Exchange",
        RecurrencePattern: { __type: "WeeklyRecurrence:#Exchange", Interval: 1, DaysOfWeek: "Monday Wednesday" },
        RecurrenceRange: { __type: "NumberedRecurrence:#Exchange", StartDate: "2026-11-02", NumberOfOccurrences: 6 },
      },
    });
  });

  it("creates all-day events with exclusive end dates", async () => {
    const { calendar, owa } = setup({
      CreateCalendarEvent: () => created("EV"),
      GetCalendarEvent: () => messages(success({ Items: [event] })),
    });
    await calendar.createEvent({ subject: "Holiday", start: "2026-12-24", isAllDay: true });
    expect(owa.of("CreateCalendarEvent")[0]!.body.Body.Items[0]).toMatchObject({
      Start: "2026-12-24T00:00:00",
      End: "2026-12-25T00:00:00",
      IsAllDayEvent: true,
      FreeBusyType: "Free",
    });
    await expect(calendar.createEvent({ subject: "x", start: "2026-12-24", end: "2026-12-24", isAllDay: true })).rejects.toThrow(/exclusive/);
  });

  it("updates with UpdateCalendarEvent (EventId + ItemChange)", async () => {
    const { calendar, owa } = setup({
      UpdateCalendarEvent: () => messages(success()),
      GetCalendarEvent: () => messages(success({ Items: [event] })),
    });
    await calendar.updateEvent("EV", { subject: "New", start: "2026-11-02T11:00:00Z", end: "2026-11-02T12:00:00Z" }, false);
    const body = owa.of("UpdateCalendarEvent")[0]!.body.Body;
    expect(body.EventId).toEqual({ __type: "ItemId:#Exchange", Id: "EV" });
    expect(body.SendCalendarInvitationsOrCancellations).toBe("SendToNone");
    expect(body.ItemChange.Updates.map((u: any) => u.Path.FieldURI)).toEqual([
      "item:Subject",
      "calendar:Start",
      "calendar:StartTimeZoneId",
      "calendar:End",
      "calendar:EndTimeZoneId",
    ]);
  });

  it("responds to invitations", async () => {
    const { calendar, owa } = setup({ CreateItem: () => messages(success()) });
    await calendar.respond("EV", "tentative", { message: "Might be late" });
    expect(owa.of("CreateItem")[0]!.body.Body).toMatchObject({
      MessageDisposition: "SendAndSaveCopy",
      Items: [{ __type: "TentativelyAcceptItem:#Exchange", ReferenceItemId: { Id: "EV" }, Body: { Value: "Might be late" } }],
    });
  });
});

describe("ContactService", () => {
  it("creates contacts with the verified dictionary-entry shapes", async () => {
    const { contacts, owa } = setup({
      CreateItem: () => created("CON"),
      GetItem: () => messages(success({ Items: [{ ItemId: { Id: "CON" }, DisplayName: "Jane Doe" }] })),
    });
    await contacts.create({ givenName: "Jane", surname: "Doe", emails: ["jane@contoso.com"], mobilePhone: "+1 555", businessAddress: { city: "Leeds" } });
    expect(owa.of("CreateItem")[0]!.body.Body.Items[0]).toEqual({
      __type: "Contact:#Exchange",
      GivenName: "Jane",
      Surname: "Doe",
      DisplayName: "Jane Doe",
      EmailAddresses: [{ Key: "EmailAddress1", EmailAddress: "jane@contoso.com", RoutingType: "SMTP" }],
      PhoneNumbers: [{ Key: "MobilePhone", PhoneNumber: "+1 555" }],
      PhysicalAddresses: [{ Key: "Business", City: "Leeds" }],
    });
  });

  it("updates indexed fields with DictionaryPropertyUri", async () => {
    const { contacts, owa } = setup({
      UpdateItem: () => messages(success()),
      GetItem: () => messages(success({ Items: [{ ItemId: { Id: "CON" } }] })),
    });
    await contacts.update("CON", { jobTitle: "Lead", emails: ["new@contoso.com"] });
    const updates = owa.of("UpdateItem")[0]!.body.Body.ItemChanges[0].Updates;
    expect(updates[0]).toMatchObject({ Path: { FieldURI: "contacts:JobTitle" }, Item: { JobTitle: "Lead" } });
    expect(updates[1]).toEqual({
      __type: "SetItemField:#Exchange",
      Path: { __type: "DictionaryPropertyUri:#Exchange", FieldURI: "contacts:EmailAddress", FieldIndex: "EmailAddress1" },
      Item: { __type: "Contact:#Exchange", EmailAddresses: [{ Key: "EmailAddress1", EmailAddress: "new@contoso.com", RoutingType: "SMTP" }] },
    });
    expect(updates[2]).toMatchObject({ __type: "DeleteItemField:#Exchange", Path: { FieldIndex: "EmailAddress2" } });
  });

  it("searches contacts and the directory, de-duplicating by address", async () => {
    const person = (name: string, email: string) => ({ DisplayName: name, EmailAddress: { Name: name, EmailAddress: email } });
    const { contacts, owa } = setup({
      FindPeople: (body) => ({
        Body: {
          ResponseClass: "Success",
          ResultSet:
            body.Body.ParentFolderId.BaseFolderId.Id === "directory"
              ? [person("Jane Doe", "jane@contoso.com"), person("Jan Smith", "jan@contoso.com")]
              : [person("Jane (personal)", "JANE@contoso.com")],
        },
      }),
    });
    const result = await contacts.searchPeople("jan");
    expect(owa.of("FindPeople").map((c) => c.body.Body.ParentFolderId.BaseFolderId)).toEqual([
      { __type: "DistinguishedFolderId:#Exchange", Id: "contacts" },
      { __type: "DistinguishedFolderId:#Exchange", Id: "directory" },
    ]);
    expect(result.people.map((p) => [p.email, p.source])).toEqual([
      ["JANE@contoso.com", "contacts"],
      ["jan@contoso.com", "directory"],
    ]);
  });
});

describe("AttachmentService", () => {
  it("downloads via GetAttachmentDownloadToken and the attachments host", async () => {
    const { attachments, owa } = setup({
      GetItem: () =>
        messages(success({ Items: [{ Attachments: [{ __type: "FileAttachment:#Exchange", AttachmentId: { Id: "A/1+" }, Name: "a.txt", ContentType: "text/plain", Size: 5 }] }] })),
      GetAttachmentDownloadToken: () => "download-jwt",
      GetFileAttachment: () => new Response("hello", { status: 200, headers: { "content-type": "text/plain" } }),
    });
    const file = await attachments.download(ID(1), { name: "A.TXT" });
    const get = owa.of("GetFileAttachment")[0]!;
    expect(get.url).toBe("https://attachments.office.net/owa/me%40contoso.com/service.svc/s/GetFileAttachment?id=A%2F1%2B");
    expect(get.headers["x-token"]).toBe("download-jwt");
    expect(file.data.toString()).toBe("hello");
  });

  it("refuses attached items and oversize files", async () => {
    const { attachments } = setup({
      GetItem: () =>
        messages(
          success({
            Items: [{ Attachments: [{ __type: "ItemAttachment:#Exchange", AttachmentId: { Id: "I" }, Name: "fwd" }, { AttachmentId: { Id: "B" }, Name: "big.bin", Size: 1e12 }] }],
          }),
        ),
    });
    await expect(attachments.download(ID(1), { attachmentId: "I" })).rejects.toThrow(/attached Outlook item/);
    await expect(attachments.download(ID(1), { name: "big.bin" })).rejects.toThrow(/limit/);
    await expect(attachments.download(ID(1), {})).rejects.toThrow(/pass attachmentId or name/);
  });
});
