import { z } from "zod";
import { defineTool, folderRef, type ToolRegistry } from "./helpers.js";

const address = z
  .object({
    street: z.string().optional(),
    city: z.string().optional(),
    state: z.string().optional(),
    postalCode: z.string().optional(),
    country: z.string().optional(),
  })
  .describe("Postal address");

const contactFields = {
  givenName: z.string().optional().describe("First name"),
  middleName: z.string().optional(),
  surname: z.string().optional().describe("Last name"),
  displayName: z.string().optional().describe('Display name (default "givenName surname")'),
  nickname: z.string().optional(),
  company: z.string().optional(),
  department: z.string().optional(),
  jobTitle: z.string().optional(),
  officeLocation: z.string().optional(),
  emails: z.array(z.string()).max(3).optional().describe('Up to three addresses ("jane@contoso.com" or "Jane Doe <jane@contoso.com>")'),
  mobilePhone: z.string().optional(),
  businessPhone: z.string().optional(),
  homePhone: z.string().optional(),
  businessAddress: address.optional(),
  homeAddress: address.optional(),
  notes: z.string().optional(),
  categories: z.array(z.string()).optional(),
};

export function registerContactTools(registry: ToolRegistry): void {
  const { contacts } = registry.services;

  defineTool(registry, {
    name: "list_contacts",
    title: "List contacts",
    kind: "read",
    description: "List personal contacts alphabetically. Use search_people to find people by name, including the organization directory.",
    input: {
      folder: folderRef("Contacts folder (default: the main Contacts folder)").optional(),
      limit: z.number().int().min(1).max(200).optional().describe("Contacts to return (default 50)"),
      offset: z.number().int().min(0).optional(),
    },
    run: (args) => contacts.list(args),
  });

  defineTool(registry, {
    name: "get_contact",
    title: "Get contact",
    kind: "read",
    description: "Get all details of a personal contact.",
    input: { id: z.string().min(1).describe("Contact ID") },
    run: ({ id }) => contacts.get(id),
  });

  defineTool(registry, {
    name: "search_people",
    title: "Search people",
    kind: "read",
    description:
      "Find people by name or email address in the organization directory (global address list) and/or personal contacts. " +
      "Useful for looking up someone's email address before sending.",
    input: {
      query: z.string().min(1).describe("Name, partial name or email address"),
      source: z.enum(["all", "directory", "contacts"]).optional().describe('Where to search (default "all")'),
      limit: z.number().int().min(1).max(100).optional().describe("Maximum results (default 10)"),
    },
    run: ({ query, source, limit }) => contacts.searchPeople(query, { source, limit }),
  });

  defineTool(registry, {
    name: "create_contact",
    title: "Create contact",
    kind: "write",
    description: "Create a personal contact.",
    input: { ...contactFields, folder: folderRef("Contacts folder (default: the main Contacts folder)").optional() },
    run: ({ folder, ...input }) => contacts.create(input, folder ?? "contacts"),
  });

  defineTool(registry, {
    name: "update_contact",
    title: "Update contact",
    kind: "write",
    description: "Update a personal contact. Only given fields change; an empty string clears a field and emails replaces all addresses.",
    input: { id: z.string().min(1).describe("Contact ID"), ...contactFields },
    run: ({ id, ...input }) => contacts.update(id, input),
  });

  defineTool(registry, {
    name: "delete_contact",
    title: "Delete contact",
    kind: "delete",
    description: "Delete a personal contact (moved to Deleted Items unless permanent).",
    input: { id: z.string().min(1).describe("Contact ID"), permanent: z.boolean().optional().describe("Permanently delete (default false)") },
    run: ({ id, permanent }) => contacts.delete(id, Boolean(permanent)),
  });
}
