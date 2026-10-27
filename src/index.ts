#!/usr/bin/env node
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { loadConfig, type Config } from "./config.js";
import { errorMessage } from "./errors.js";
import { createServer, createServices, SERVER_NAME, SERVER_VERSION } from "./server.js";

const HELP = `${SERVER_NAME} ${SERVER_VERSION}
MCP server (stdio) for Microsoft 365 Outlook via the Outlook on the web service API.

Usage:
  outlook-mcp            Start the MCP server on stdio
  outlook-mcp --check    Verify authentication and print the mailbox identity
  outlook-mcp --help     Show this help
  outlook-mcp --version  Print the version

Authentication (first match wins unless OUTLOOK_AUTH is set):
  OUTLOOK_ACCESS_TOKEN    Bearer token for https://outlook.office.com
  OUTLOOK_TOKEN_FILE      File containing the token (re-read when it expires)
  OUTLOOK_TOKEN_COMMAND   Shell command that prints a token
  OUTLOOK_COOKIE          OWA session cookie (+ OUTLOOK_CANARY), experimental
  (default)               Azure CLI: az account get-access-token --resource https://outlook.office.com

Other settings:
  OUTLOOK_AUTH                 auto | azure-cli | token | token-file | command | cookie
  OUTLOOK_AZ_TENANT            Tenant for Azure CLI token requests
  OUTLOOK_READ_ONLY=1          Only register read-only tools
  OUTLOOK_TIMEZONE             Windows time zone ID for local times (default: mailbox setting)
  OUTLOOK_ORIGIN               Default https://outlook.cloud.microsoft
  OUTLOOK_ANCHOR_MAILBOX       Override X-AnchorMailbox
  OUTLOOK_MAX_DOWNLOAD_MB      Attachment size limit (default 25)
  OUTLOOK_REQUEST_TIMEOUT_MS   Default 30000
  OUTLOOK_DEBUG=1              Log requests to stderr (tokens are never logged)
`;

const log = (message: string) => process.stderr.write(`[${SERVER_NAME}] ${message}\n`);

async function check(config: Config): Promise<void> {
  const services = createServices(config, { log });
  const info = await services.mailbox.info();
  const inbox = (await services.folders.all()).find((f) => f.wellKnownName === "inbox");
  process.stdout.write(
    JSON.stringify(
      {
        ok: true,
        auth: config.auth.kind,
        mailbox: info.email,
        displayName: info.displayName,
        timeZone: info.timeZone,
        inbox: inbox ? { total: inbox.totalCount, unread: inbox.unreadCount } : undefined,
        readOnly: config.readOnly,
      },
      null,
      2,
    ) + "\n",
  );
}

async function main(): Promise<void> {
  const args = new Set(process.argv.slice(2));
  if (args.has("--help") || args.has("-h")) {
    process.stdout.write(HELP);
    return;
  }
  if (args.has("--version") || args.has("-v")) {
    process.stdout.write(`${SERVER_VERSION}\n`);
    return;
  }
  const config = loadConfig();
  if (args.has("--check")) {
    await check(config);
    return;
  }

  const { server, tools } = createServer(config, { log });
  const transport = new StdioServerTransport();
  await server.connect(transport);
  log(`ready: ${tools.length} tools, auth=${config.auth.kind}${config.readOnly ? ", read-only" : ""}`);

  const shutdown = async () => {
    await server.close().catch(() => {});
    process.exit(0);
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}

main().catch((error) => {
  log(`fatal: ${errorMessage(error)}`);
  process.exit(1);
});
