import { defineTool, type ToolRegistry } from "./helpers.js";

export function registerMailboxTools(registry: ToolRegistry): void {
  const { mailbox, folders } = registry.services;

  defineTool(registry, {
    name: "get_mailbox_info",
    title: "Mailbox info",
    kind: "read",
    description:
      "Get the signed-in user's name, email address, time zone (Windows ID, used for times given without an offset), storage quota, and inbox/drafts counts.",
    input: {},
    run: async () => {
      const [info, all] = await Promise.all([mailbox.info(), folders.all()]);
      const counts = Object.fromEntries(
        all
          .filter((f) => f.wellKnownName && ["inbox", "drafts", "sentitems", "deleteditems", "junkemail", "archive"].includes(f.wellKnownName))
          .map((f) => [f.wellKnownName!, { total: f.totalCount ?? 0, unread: f.unreadCount ?? 0 }]),
      );
      return { ...info, folders: counts, currentTimeUtc: new Date().toISOString() };
    },
  });

  defineTool(registry, {
    name: "list_categories",
    title: "List categories",
    kind: "read",
    description: "List the mailbox's category names and colors (for categorize_messages and events).",
    input: {},
    run: async () => ({ categories: await mailbox.categories() }),
  });
}
