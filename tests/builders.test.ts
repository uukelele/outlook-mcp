import { describe, expect, it } from "vitest";
import {
  attendees,
  distinguishedFolder,
  envelope,
  folderIdRef,
  looksLikeExchangeId,
  mailRecipients,
  parseAddress,
  typeKeysFirst,
} from "../src/owa/builders.js";
import { ID } from "./fake-owa.js";

describe("envelope", () => {
  it("wraps the body with JsonRequest/Request types and a UTC time zone header", () => {
    const env = envelope("FindItem", { Traversal: "Shallow" });
    expect(env).toEqual({
      __type: "FindItemJsonRequest:#Exchange",
      Header: {
        __type: "JsonRequestHeaders:#Exchange",
        RequestServerVersion: "V2018_01_08",
        TimeZoneContext: {
          __type: "TimeZoneContext:#Exchange",
          TimeZoneDefinition: { __type: "TimeZoneDefinitionType:#Exchange", Id: "UTC" },
        },
      },
      Body: { __type: "FindItemRequest:#Exchange", Traversal: "Shallow" },
    });
  });

  it("uses jsonType for actions whose envelope type differs (CreateCalendarEvent)", () => {
    const env = envelope("CreateCalendarEvent", {}, { jsonType: "CreateItem", timeZone: "GMT Standard Time" });
    expect(env.__type).toBe("CreateItemJsonRequest:#Exchange");
    expect((env.Body as any).__type).toBe("CreateItemRequest:#Exchange");
    expect((env.Header as any).TimeZoneContext.TimeZoneDefinition.Id).toBe("GMT Standard Time");
  });

  it("keeps __type first at every level, as the server's deserializer requires", () => {
    const env = envelope("CreateItem", { Items: [{ __type: "Message:#Exchange", Subject: "x" }] });
    expect(JSON.stringify(env).startsWith('{"__type"')).toBe(true);
    expect(typeKeysFirst(env)).toBe(true);
    expect(typeKeysFirst({ a: 1, __type: "X" })).toBe(false);
  });
});

describe("folder references", () => {
  it("maps well-known names and aliases to distinguished IDs", () => {
    expect(distinguishedFolder("Inbox")).toBe("inbox");
    expect(distinguishedFolder("sent")).toBe("sentitems");
    expect(distinguishedFolder("Deleted Items")).toBe("deleteditems");
    expect(distinguishedFolder("trash")).toBe("deleteditems");
    expect(distinguishedFolder("spam")).toBe("junkemail");
    expect(distinguishedFolder("Projects")).toBeUndefined();
  });

  it("builds DistinguishedFolderId or FolderId", () => {
    expect(folderIdRef("inbox")).toEqual({ __type: "DistinguishedFolderId:#Exchange", Id: "inbox" });
    expect(folderIdRef(ID(1))).toEqual({ __type: "FolderId:#Exchange", Id: ID(1) });
  });

  it("recognizes Exchange IDs", () => {
    expect(looksLikeExchangeId(ID(1))).toBe(true);
    expect(looksLikeExchangeId("Inbox/Projects")).toBe(false);
  });
});

describe("addresses", () => {
  it("parses plain and named addresses", () => {
    expect(parseAddress("jane@contoso.com")).toEqual({ email: "jane@contoso.com" });
    expect(parseAddress("Jane Doe <jane@contoso.com>")).toEqual({ email: "jane@contoso.com", name: "Jane Doe" });
    expect(parseAddress('"Doe, Jane" <jane@contoso.com>')).toEqual({ email: "jane@contoso.com", name: "Doe, Jane" });
    expect(() => parseAddress("not an address")).toThrow(/not a valid email/);
  });

  it("builds flat mail recipients without __type (nested Mailbox does not resolve in OWA JSON)", () => {
    expect(mailRecipients(["Jane <jane@contoso.com>", "bob@contoso.com"])).toEqual([
      { Name: "Jane", EmailAddress: "jane@contoso.com", RoutingType: "SMTP" },
      { EmailAddress: "bob@contoso.com", RoutingType: "SMTP" },
    ]);
  });

  it("wraps calendar attendees in Mailbox", () => {
    expect(attendees(["bob@contoso.com"])).toEqual([
      { __type: "AttendeeType:#Exchange", Mailbox: { EmailAddress: "bob@contoso.com", RoutingType: "SMTP" } },
    ]);
  });
});
