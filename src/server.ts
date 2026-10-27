import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createAuthProvider, type AuthProvider } from "./auth/providers.js";
import type { Config } from "./config.js";
import { OwaClient, type FetchLike } from "./owa/client.js";
import { AttachmentService } from "./services/attachments.js";
import { CalendarService } from "./services/calendar.js";
import { ContactService } from "./services/contacts.js";
import { FolderService } from "./services/folders.js";
import { MailService } from "./services/mail.js";
import { MailboxService } from "./services/mailbox.js";
import { registerCalendarTools } from "./tools/calendar.js";
import { registerContactTools } from "./tools/contacts.js";
import type { Services, ToolRegistry } from "./tools/helpers.js";
import { registerMailboxTools } from "./tools/mailbox.js";
import { registerMailTools } from "./tools/mail.js";

export const SERVER_NAME = "outlook-mcp";
export const SERVER_VERSION = "0.1.0";

export interface ServerOptions {
  log?: (message: string) => void;
  /** Override auth (tests). */
  auth?: AuthProvider;
  /** Override fetch (tests). */
  fetch?: FetchLike;
}

export function createServices(config: Config, options: ServerOptions = {}): Services & { client: OwaClient } {
  const log = options.log ?? (() => {});
  const client = new OwaClient({
    origin: config.origin,
    attachmentsOrigin: config.attachmentsOrigin,
    auth: options.auth ?? createAuthProvider(config.auth, log),
    anchorMailbox: config.anchorMailbox,
    userAgent: config.userAgent,
    requestTimeoutMs: config.requestTimeoutMs,
    fetch: options.fetch,
    log,
    debug: config.debug,
  });
  const mailbox = new MailboxService(client, config.timeZone);
  const folders = new FolderService(client);
  const attachments = new AttachmentService(client, mailbox, config.maxDownloadBytes);
  return {
    client,
    mailbox,
    folders,
    attachments,
    mail: new MailService(client, folders, attachments),
    calendar: new CalendarService(client, folders, mailbox),
    contacts: new ContactService(client, folders),
  };
}

export function createServer(config: Config, options: ServerOptions = {}): { server: McpServer; tools: string[]; services: Services } {
  const services = createServices(config, options);
  const server = new McpServer(
    { name: SERVER_NAME, version: SERVER_VERSION },
    {
      instructions:
        "Tools for the user's Microsoft 365 Outlook mailbox (mail, calendar, contacts) via Outlook on the web. " +
        "IDs returned by one tool can be passed to others. Times in results are UTC; call get_mailbox_info for the user's time zone. " +
        "send_message, reply_message, forward_message, send_draft, create_event with attendees, update_event, delete_event and respond_to_event " +
        "contact other people — confirm with the user before using them. Prefer create_draft when the user has not explicitly asked to send." +
        (config.readOnly ? " This server is running in read-only mode; no changes can be made." : ""),
    },
  );
  const registry: ToolRegistry = { server, services, readOnly: config.readOnly, registered: [] };
  registerMailboxTools(registry);
  registerMailTools(registry);
  registerCalendarTools(registry);
  registerContactTools(registry);
  return { server, tools: registry.registered, services };
}
