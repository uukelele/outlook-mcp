import { OutlookError } from "../errors.js";
import {
  bodyContent,
  chunk,
  deleteItemField,
  folderIdRef,
  indexedPage,
  indexedPropertyUri,
  itemChange,
  itemId,
  itemShape,
  parseAddress,
  propertyUri,
  setItemField,
  sortBy,
  t,
  targetFolder,
  type Json,
} from "../owa/builders.js";
import { assertBodySuccess, responseMessages, type OwaClient } from "../owa/client.js";
import { contactDetail, contactSummary, persona } from "../owa/normalize.js";
import type { FolderService } from "./folders.js";

export interface PostalAddress {
  street?: string;
  city?: string;
  state?: string;
  postalCode?: string;
  country?: string;
}

export interface ContactInput {
  givenName?: string;
  middleName?: string;
  surname?: string;
  displayName?: string;
  nickname?: string;
  company?: string;
  department?: string;
  jobTitle?: string;
  officeLocation?: string;
  /** Up to three addresses; "Name <email>" or plain address. */
  emails?: string[];
  mobilePhone?: string;
  businessPhone?: string;
  homePhone?: string;
  businessAddress?: PostalAddress;
  homeAddress?: PostalAddress;
  notes?: string;
  categories?: string[];
}

export type PeopleSource = "directory" | "contacts" | "all";

const EMAIL_KEYS = ["EmailAddress1", "EmailAddress2", "EmailAddress3"];
const PHONE_FIELDS = [
  ["mobilePhone", "MobilePhone"],
  ["businessPhone", "BusinessPhone"],
  ["homePhone", "HomePhone"],
] as const;
const ADDRESS_FIELDS = [
  ["businessAddress", "Business"],
  ["homeAddress", "Home"],
] as const;
const ADDRESS_PARTS = [
  ["street", "Street"],
  ["city", "City"],
  ["state", "State"],
  ["postalCode", "PostalCode"],
  ["country", "CountryOrRegion"],
] as const;

export class ContactService {
  constructor(
    private readonly client: OwaClient,
    private readonly folders: FolderService,
  ) {}

  async list(options: { folder?: string; limit?: number; offset?: number } = {}) {
    const folderId = await this.folders.resolve(options.folder ?? "contacts");
    const limit = Math.max(1, Math.min(options.limit ?? 50, 200));
    const offset = Math.max(0, options.offset ?? 0);
    const [message] = await this.client.callForMessages("FindItem", {
      ItemShape: itemShape(["contacts:DisplayName"]),
      ParentFolderIds: [folderIdRef(folderId)],
      Traversal: "Shallow",
      Paging: indexedPage(offset, limit),
      SortOrder: sortBy("contacts:DisplayName", "Ascending"),
    });
    const root = (message?.RootFolder as Json | undefined) ?? {};
    const ids = ((root.Items as Json[] | undefined) ?? []).map((item) => (item.ItemId as Json).Id as string);
    const contacts = await this.fetchMany(ids);
    return {
      total: root.TotalItemsInView,
      offset,
      count: contacts.length,
      nextOffset: root.IncludesLastItemInRange === false ? offset + ids.length : undefined,
      contacts: contacts.map(contactSummary),
    };
  }

  private async fetchMany(ids: string[]): Promise<Json[]> {
    const out: Json[] = [];
    for (const group of chunk(ids, 50)) {
      const result = await this.client.call("GetItem", {
        ItemShape: { __type: t("ItemResponseShape"), BaseShape: "AllProperties", BodyType: "Text" },
        ItemIds: group.map((id) => itemId(id)),
      });
      for (const message of responseMessages(result)) {
        const item = (message.Items as Json[] | undefined)?.[0];
        if (item) out.push(item);
      }
    }
    return out;
  }

  async get(id: string) {
    const [item] = await this.fetchMany([id]);
    if (!item) throw new OutlookError("Contact not found");
    return contactDetail(item);
  }

  async create(input: ContactInput, folder = "contacts") {
    if (!input.displayName && !input.givenName && !input.surname && !input.company && !input.emails?.length) {
      throw new OutlookError("A contact needs at least a name, company or email address");
    }
    if ((input.emails?.length ?? 0) > 3) throw new OutlookError("Outlook contacts hold at most three email addresses");
    const contact: Json = { __type: t("Contact") };
    const displayName = input.displayName ?? ([input.givenName, input.surname].filter(Boolean).join(" ") || undefined);
    const scalar: [string, unknown][] = [
      ["GivenName", input.givenName],
      ["MiddleName", input.middleName],
      ["Surname", input.surname],
      ["DisplayName", displayName],
      ["Nickname", input.nickname],
      ["CompanyName", input.company],
      ["Department", input.department],
      ["JobTitle", input.jobTitle],
      ["OfficeLocation", input.officeLocation],
    ];
    for (const [key, value] of scalar) if (value) contact[key] = value;
    if (input.emails?.length) contact.EmailAddresses = input.emails.map((email, i) => emailEntry(EMAIL_KEYS[i]!, email));
    const phones = PHONE_FIELDS.filter(([field]) => input[field]).map(([field, key]) => ({ Key: key, PhoneNumber: input[field] }));
    if (phones.length) contact.PhoneNumbers = phones;
    const addresses = ADDRESS_FIELDS.filter(([field]) => input[field]).map(([field, key]) => addressEntry(key, input[field]!));
    if (addresses.length) contact.PhysicalAddresses = addresses;
    if (input.notes) contact.Body = bodyContent(input.notes, "text");
    if (input.categories?.length) contact.Categories = input.categories;

    const folderId = await this.folders.resolve(folder);
    const [message] = await this.client.callForMessages("CreateItem", {
      SavedItemFolderId: targetFolder(folderId),
      Items: [contact],
    });
    const id = (((message?.Items as Json[] | undefined)?.[0]?.ItemId as Json | undefined)?.Id as string | undefined) ?? undefined;
    if (!id) throw new OutlookError("CreateItem succeeded but returned no contact ID");
    return this.get(id);
  }

  /** Update fields; pass an empty string to clear a text field, or an empty array to clear emails/categories. */
  async update(id: string, input: ContactInput) {
    const updates: Json[] = [];
    const scalar = (uri: string, key: string, value: string | undefined) => {
      if (value === undefined) return;
      updates.push(value === "" ? deleteItemField(propertyUri(uri)) : setItemField(propertyUri(uri), "Contact", { [key]: value }));
    };
    scalar("contacts:GivenName", "GivenName", input.givenName);
    scalar("contacts:MiddleName", "MiddleName", input.middleName);
    scalar("contacts:Surname", "Surname", input.surname);
    scalar("contacts:DisplayName", "DisplayName", input.displayName);
    scalar("contacts:Nickname", "Nickname", input.nickname);
    scalar("contacts:CompanyName", "CompanyName", input.company);
    scalar("contacts:Department", "Department", input.department);
    scalar("contacts:JobTitle", "JobTitle", input.jobTitle);
    scalar("contacts:OfficeLocation", "OfficeLocation", input.officeLocation);

    if (input.emails !== undefined) {
      if (input.emails.length > 3) throw new OutlookError("Outlook contacts hold at most three email addresses");
      EMAIL_KEYS.forEach((key, i) => {
        const path = indexedPropertyUri("contacts:EmailAddress", key);
        const email = input.emails![i];
        updates.push(email ? setItemField(path, "Contact", { EmailAddresses: [emailEntry(key, email)] }) : deleteItemField(path));
      });
    }
    for (const [field, key] of PHONE_FIELDS) {
      const value = input[field];
      if (value === undefined) continue;
      const path = indexedPropertyUri("contacts:PhoneNumber", key);
      updates.push(value ? setItemField(path, "Contact", { PhoneNumbers: [{ Key: key, PhoneNumber: value }] }) : deleteItemField(path));
    }
    for (const [field, key] of ADDRESS_FIELDS) {
      const address = input[field];
      if (!address) continue;
      for (const [part, partKey] of ADDRESS_PARTS) {
        const value = address[part];
        if (value === undefined) continue;
        const path = indexedPropertyUri(`contacts:PhysicalAddress:${partKey}`, key);
        updates.push(
          value ? setItemField(path, "Contact", { PhysicalAddresses: [{ Key: key, [partKey]: value }] }) : deleteItemField(path),
        );
      }
    }
    if (input.notes !== undefined) {
      updates.push(input.notes ? setItemField(propertyUri("item:Body"), "Contact", { Body: bodyContent(input.notes, "text") }) : deleteItemField(propertyUri("item:Body")));
    }
    if (input.categories !== undefined) {
      updates.push(
        input.categories.length
          ? setItemField(propertyUri("item:Categories"), "Contact", { Categories: input.categories })
          : deleteItemField(propertyUri("item:Categories")),
      );
    }
    if (!updates.length) throw new OutlookError("Nothing to update");
    await this.client.callForMessages("UpdateItem", { ConflictResolution: "AlwaysOverwrite", ItemChanges: [itemChange(id, updates)] });
    return this.get(id);
  }

  async delete(id: string, permanent = false) {
    await this.client.callForMessages("DeleteItem", {
      ItemIds: [itemId(id)],
      DeleteType: permanent ? "HardDelete" : "MoveToDeletedItems",
    });
    return { deleted: true };
  }

  /** Search the organization directory (GAL) and/or personal contacts via FindPeople. */
  async searchPeople(query: string, options: { source?: PeopleSource; limit?: number } = {}) {
    const text = query.trim();
    if (!text) throw new OutlookError("query must not be empty");
    const limit = Math.max(1, Math.min(options.limit ?? 10, 100));
    const source = options.source ?? "all";
    const sources: ("directory" | "contacts")[] = source === "all" ? ["contacts", "directory"] : [source];
    const seen = new Set<string>();
    const results: (ReturnType<typeof persona> & { source: string })[] = [];
    for (const folder of sources) {
      const result = await this.client.call(
        "FindPeople",
        {
          IndexedPageItemView: indexedPage(0, limit),
          ParentFolderId: targetFolder(folder === "directory" ? "directory" : "contacts"),
          PersonaShape: { __type: t("PersonaResponseShape"), BaseShape: "Default" },
          QueryString: text,
          SearchPeopleSuggestionIndex: false,
          ShouldResolveOneOffEmailAddress: false,
        },
      );
      const body = assertBodySuccess("FindPeople", result);
      for (const raw of (body.ResultSet as Json[] | undefined) ?? []) {
        const entry = persona(raw);
        const key = (entry.email ?? entry.displayName ?? "").toLowerCase();
        if (key && seen.has(key)) continue;
        if (key) seen.add(key);
        results.push({ ...entry, source: folder });
      }
    }
    return { query: text, count: Math.min(results.length, limit), people: results.slice(0, limit) };
  }
}

function emailEntry(key: string, input: string): Json {
  const { email, name } = parseAddress(input);
  return name ? { Key: key, Name: name, EmailAddress: email, RoutingType: "SMTP" } : { Key: key, EmailAddress: email, RoutingType: "SMTP" };
}

function addressEntry(key: string, address: PostalAddress): Json {
  const entry: Json = { Key: key };
  for (const [part, partKey] of ADDRESS_PARTS) if (address[part]) entry[partKey] = address[part];
  return entry;
}
