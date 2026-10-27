# Outlook on the web (OWA) service API — protocol notes

These notes document the internal JSON API used by Outlook on the web, as used by this server. Everything marked **verified** was observed working against a live Microsoft 365 tenant (October 2026) with an Azure CLI token; items marked **reference** come from production open-source clients and the EWS schema and are covered by payload unit tests, but were not exercised live.

Prior art that informed this work: [mustang-im/mustang](https://github.com/mustang-im/mustang) (TypeScript, OWA accounts), [jlentink/outlook-web-mcp](https://github.com/jlentink/outlook-web-mcp) (Go), [ma-nakaya/outlook-web-mcp-server](https://github.com/ma-nakaya/outlook-web-mcp-server), [nhype/owa-exchange-mcp](https://github.com/nhype/owa-exchange-mcp), [ilyabazhenov/mac-owa-widget](https://github.com/ilyabazhenov/mac-owa-widget), [johnkil/outlook-agent](https://github.com/johnkil/outlook-agent).

## Transport

```
POST https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=Mail&n=<counter>
```

| Header | Value |
| --- | --- |
| `Authorization` | `Bearer <token>` — audience `https://outlook.office.com`; the Azure CLI first-party token (`scp=user_impersonation`) works |
| `Action` | the action name (required, in addition to the query string) |
| `Content-Type` | `application/json; charset=utf-8` |
| `Prefer` | `IdType="ImmutableId", exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"` — immutable IDs survive moves |
| `X-AnchorMailbox` | `PUID:<puid>@<tenant-id>` from the token claims (verified), or the SMTP address (verified) |
| `X-Requested-With`, `Origin`, `Referer`, `X-OWA-CorrelationId`, `X-OWA-SessionId`, `client-request-id` | as the web app sends them |

The JSON payload goes in the body. (OWA also accepts it URL-encoded in an `X-OWA-UrlPostData` header with an empty body — verified for `GetCalendarView` and `GetOwaUserConfiguration` — but the body works for everything and has no size limit.)

`n` is a cache-busting counter. `app=Mail` works for calendar and people actions too.

Redirects are never followed (a redirect means a bounce to sign-in, and following it would leak the token to another host). `401`/`440` → refresh the token once and retry. `429`/`503` → back off and retry.

## Envelope

```json
{
  "__type": "<Action>JsonRequest:#Exchange",
  "Header": {
    "__type": "JsonRequestHeaders:#Exchange",
    "RequestServerVersion": "V2018_01_08",
    "TimeZoneContext": {
      "__type": "TimeZoneContext:#Exchange",
      "TimeZoneDefinition": { "__type": "TimeZoneDefinitionType:#Exchange", "Id": "UTC" }
    }
  },
  "Body": { "__type": "<Action>Request:#Exchange", "...": "..." }
}
```

- `__type` must be the **first key** of every polymorphic object (the server uses .NET `DataContractJsonSerializer`).
- The time zone context is a Windows zone ID. With `UTC`, every returned time ends in `Z`. Times sent *without* an offset are interpreted in this zone — verified: `10:00` in `GMT Standard Time` on a June date was stored as `09:00Z`.
- `CreateCalendarEvent` uses the `CreateItem` envelope types (`CreateItemJsonRequest` / `CreateItemRequest`) with the `CreateCalendarEvent` action.
- `GetOwaUserConfiguration` and `GetAttachmentDownloadToken` take their own payloads (below), not the standard body.

## Responses and errors

| Action | Results at |
| --- | --- |
| most actions | `Body.ResponseMessages.Items[]`, one per requested item, in request order |
| `FindItem` / `FindFolder` | `…Items[0].RootFolder.Items` / `.Folders`, plus `TotalItemsInView`, `IncludesLastItemInRange` |
| `GetCalendarView` | `Body.Items` (no `ResponseMessages`) |
| `FindPeople` | `Body.ResultSet` |
| `GetConversationItems` | `…Items[0].Conversation.ConversationNodes[].Items[]` |
| `GetOwaUserConfiguration` | top level: `SessionSettings`, `UserOptions`, `MasterCategoryList`, … |
| `GetAttachmentDownloadToken` | a bare JSON string |

Errors come in two forms:

- **HTTP 200 with `ResponseClass: "Error"`** and a `ResponseCode`/`MessageText` per item (e.g. `ErrorItemNotFound`). Batch operations report these per item.
- **HTTP 4xx/5xx** with `Body.ExceptionName` and `FaultMessage`, e.g. `OwaMethodArgumentException` (request shape rejected), or `… does not exist in the service method map` (action not available on this endpoint).

## Actions

### Mail — verified

| Purpose | Action | Key details |
| --- | --- | --- |
| Folder tree | `FindFolder` | `Traversal: "Deep"` from `msgfolderroot`; properties `folder:DisplayName`, `folder:ParentFolderId`, `folder:FolderClass`, `folder:TotalCount`, `folder:UnreadCount`, `folder:ChildFolderCount`, `folder:DistinguishedFolderId` |
| Create / rename / delete folder | `CreateFolder`, `UpdateFolder` (`SetFolderField` on `folder:DisplayName`), `DeleteFolder` | |
| List messages | `FindItem` | `IndexedPageView` paging, `SortOrder` on `item:DateTimeReceived`, `ViewFilter` = `All`/`Unread`/`Flagged`/`ToOrCcMe` (`HasAttachments` is rejected) |
| Read message | `GetItem` | `BaseShape: "IdOnly"` + properties; `BodyType: "Text"` returns server-converted text; `"HTML"` + `FilterHtmlContent: true` returns sanitized HTML |
| Conversation | `GetConversationItems` | `ShapeName: "ItemPart"`, `SortOrder: "DateOrderDescending"`; `item:UniqueBody` is each message's new text only |
| Send new mail | `CreateItem` | `MessageDisposition: "SendAndSaveCopy"`, `SavedItemFolderId` = `sentitems` |
| Draft | `CreateItem` | `MessageDisposition: "SaveOnly"`, `SavedItemFolderId` = `drafts` |
| Edit draft | `UpdateItem` | `SetItemField` on `item:Subject`, `item:Body` |
| Send draft | `UpdateItem` | `MessageDisposition: "SendAndSaveCopy"` with at least one `SetItemField` (the subject is re-set). There is **no `SendItem`** action |
| Reply / reply-all / forward | `CreateItem` | items `ReplyToItem` / `ReplyAllToItem` / `ForwardItem` with `ReferenceItemId`, `NewBodyContent` and either the original's `ChangeKey` or `ShouldIgnoreChangeKey: true` (one is required). Non-mail items such as delivery reports cannot be replied to |
| Read / flag / categories | `UpdateItem` | paths `IsRead`, `item:Flag` (`{"__type":"FlagType:#Exchange","FlagStatus":"Flagged"\|"Complete"\|"NotFlagged"}`), `item:Categories` (string array); `SuppressReadReceipts: true` |
| Move / copy | `MoveItem` / `CopyItem` | `ToFolderId` + `ReturnNewItemIds`; with immutable IDs a moved item keeps its ID |
| Delete | `DeleteItem` | `DeleteType` `MoveToDeletedItems` or `HardDelete` |

**Recipients are flat objects without `__type`:**

```json
"ToRecipients": [{ "Name": "Jane Doe", "EmailAddress": "jane@contoso.com", "RoutingType": "SMTP" }]
```

The EWS-style nested form (`{"__type":"EmailAddressWrapper:#Exchange","Mailbox":{…}}`) is accepted but the recipient does not resolve and sending fails with "At least one recipient is not valid".

Returned addresses can be Exchange legacy DNs (`RoutingType: "EX"`, `/O=EXCHANGELABS/…`) with the SMTP address in `SmtpEmailAddress`.

### Attachments — verified

| Purpose | Request |
| --- | --- |
| List | `GetItem` with property `item:Attachments` (`FileAttachment`, `ItemAttachment`, `ReferenceAttachment`) |
| Upload | `CreateAttachment` — `ParentItemId`, `Attachments: [{"__type":"FileAttachment:#Exchange","Name","ContentType","IsInline","Content":<base64>}]` |
| Remove | `DeleteAttachment` — `AttachmentIds: [{"__type":"AttachmentId:#Exchange","Id"}]` |
| Download | 1. `POST …?action=GetAttachmentDownloadToken` with raw body `{"__type":"GetAttachmentDownloadTokenRequest:#Exchange","UseNewDownloadTokenLifetime":true}` → JWT string<br>2. `GET https://attachments.office.net/owa/<smtp>/service.svc/s/GetFileAttachment?id=<attachment id>` with header `x-token: <jwt>` |

`GetAttachment` (HTTP 400) and `service.svc/s/GetFileAttachment` with a bearer token (HTTP 500) do not work on the cloud endpoint.

### Calendar — verified

| Purpose | Action | Key details |
| --- | --- | --- |
| Events in range | `GetCalendarView` | `CalendarId` (`TargetFolderId`), `RangeStart`, `RangeEnd`; recurring series are expanded. (`FindItem` with `CalendarView` paging is silently ignored; `CalendarPageView` works but is redundant) |
| Event details | `GetCalendarEvent` | `EventIds`, `ItemShape` (`IdOnly` + `BodyType`). Returns attendees, organizer, body, time zone — much richer than `GetItem` |
| Create | `CreateCalendarEvent` | `CreateItem` envelope; `SavedItemFolderId`, `SendMeetingInvitations`, `Items: [CalendarItem]` with `Start`, `End`, `StartTimeZoneId`, `EndTimeZoneId`, `Location` (`EnhancedLocation`), attendees |
| Update | `UpdateCalendarEvent` | `EventId`, `ItemChange { ItemId, Updates[] }`, `SendCalendarInvitationsOrCancellations`; verified for `item:Subject`, `calendar:Start`, `calendar:End` |
| Delete | `DeleteItem` | with `SendMeetingCancellations` |

**Attendees are wrapped** (unlike mail recipients):

```json
"RequiredAttendees": [{ "__type": "AttendeeType:#Exchange", "Mailbox": { "EmailAddress": "bob@contoso.com", "Name": "Bob", "RoutingType": "SMTP" } }]
```

**`SendMeetingInvitations: "SendToNone"` is ignored when creating an event with attendees** — invitations were sent anyway. Treat any create with attendees as sending invitations.

Reference-only (from Mustang's Office 365 code path): accept/tentative/decline via `CreateItem` with `AcceptItem` / `TentativelyAcceptItem` / `DeclineItem` and `ReferenceItemId`; recurrence (`RecurrenceType` with `DailyRecurrence` / `WeeklyRecurrence` / `AbsoluteMonthlyRecurrence` / `AbsoluteYearlyRecurrence` and `NoEnd` / `EndDate` / `Numbered` ranges); Teams meetings (`IsOnlineMeeting` + `OnlineMeetingProvider: "TeamsForBusiness"`); updating `EnhancedLocation`, attendees, reminders.

### Contacts and people — verified

| Purpose | Request |
| --- | --- |
| List | `FindItem` on `contacts` (sort `contacts:DisplayName`), then batched `GetItem` with `BaseShape: "AllProperties"` (`FindItem` does not return email addresses) |
| Create | `CreateItem` with `Contact`: `EmailAddresses: [{"Key":"EmailAddress1","EmailAddress","Name","RoutingType":"SMTP"}]`, `PhoneNumbers: [{"Key":"MobilePhone","PhoneNumber"}]`, `PhysicalAddresses: [{"Key":"Business","Street","City","PostalCode","CountryOrRegion"}]` |
| Update | `UpdateItem`; indexed fields use `{"__type":"DictionaryPropertyUri:#Exchange","FieldURI":"contacts:EmailAddress","FieldIndex":"EmailAddress1"}` with the item carrying a one-entry array |
| Search | `FindPeople` with `ParentFolderId` = distinguished `directory` (organization address list) or `contacts`, `QueryString`, `PersonaShape` |

Email dictionary entries must use `EmailAddress` (not `Value`, which fails with "EmailAddress or ItemId must be included"). `ResolveNames` **does not exist** on the cloud endpoint ("does not exist in the service method map").

### Mailbox

`GetOwaUserConfiguration` (raw payload `{"__type":"GetOwaUserConfigurationJsonRequest:#Exchange","Header":{…},"Body":{"__type":"GetOwaUserConfigurationRequest:#Exchange","Owaconfigs":["MasterCategoryList"]}}`) returns `SessionSettings` (display name, SMTP address, UPN, quota, max message size), `UserOptions.TimeZone` (Windows ID) and `MasterCategoryList.MasterList` (categories with preset color indexes).

## Search

No server-side search path works with Azure CLI tokens:

| Attempt | Result |
| --- | --- |
| `FindItem` `QueryString` as a string | HTTP 400 `OwaMethodArgumentException` |
| `FindItem` `QueryString` as `QueryStringType` | accepted, but intermittently fails creating its "EWS Search" folder and otherwise returns no results |
| `FindConversation` + `QueryString` | no results |
| `FindItem` `Restriction` (`IsEqualTo`, `Contains` on subject/body/from/IsRead) | `ErrorUnsupportedPathForQuery` |
| `outlook.cloud.microsoft/searchservice/api/v2/query`, `substrate.office.com/search/api/v2/query` | HTTP 401 with both Outlook- and Substrate-audience tokens (needs scopes only Outlook's own SPA client receives) |

`search_messages` therefore scans folder listings (subject, sender, recipients' display names, preview) client-side.

## Verification record

All write verification was done on throwaway items (drafts, a test folder, test events, a test contact, mail addressed to the account itself), each deleted immediately afterwards. Sending to the account's own address produced "Undeliverable" reports on the tenant used for verification even though the messages were accepted and saved in Sent Items — a tenant transport policy, not an API problem. No mail was sent to other people, so external delivery was not observed directly. Everything after initial development is checked with offline payload tests and the read-only live suite (`npm run test:live`).

| Area | Status |
| --- | --- |
| Folders: list, get, create, rename, hard delete | verified |
| Folders: move, delete to Deleted Items | reference |
| Messages: list, filter, sort, page, read (text/HTML), thread | verified |
| Send new, draft create/edit/send, reply/forward drafts | verified |
| Update draft recipients, forward cc/bcc, reply sent directly, send with attachments end-to-end | reference (component calls verified) |
| Read state, flags, categories, move, copy, delete | verified |
| Attachments: list, upload, download, delete | verified |
| Calendar: view, details, create (with and without attendees), update subject/time, delete | verified |
| Calendar: invitation responses, recurrence, all-day, Teams link, location/attendee/reminder updates | reference |
| Contacts: create, read, update (incl. indexed email/phone), delete; directory + contacts search | verified |
| Contacts: postal address updates | reference |
| Cookie + canary authentication | reference (untested) |
