import { OutlookError } from "../errors.js";
import { distinguishedFolder, folderIdRef, looksLikeExchangeId, propertyUri, t, targetFolder, type Json } from "../owa/builders.js";
import type { OwaClient } from "../owa/client.js";
import { folder as normalizeFolder } from "../owa/normalize.js";

export type FolderKind = "mail" | "calendar" | "contacts" | "tasks" | "all";

export interface FolderInfo {
  id: string;
  name?: string;
  path: string;
  wellKnownName?: string;
  parentId?: string;
  folderClass?: string;
  totalCount?: number;
  unreadCount?: number;
  childFolderCount?: number;
}

const FOLDER_PROPS = [
  "folder:DisplayName",
  "folder:ParentFolderId",
  "folder:FolderClass",
  "folder:TotalCount",
  "folder:UnreadCount",
  "folder:ChildFolderCount",
  "folder:DistinguishedFolderId",
];

const CACHE_TTL_MS = 5 * 60 * 1000;

/** Well-known system folders that are not useful to show in mail folder listings. */
const HIDDEN_WELL_KNOWN = new Set([
  "conflicts",
  "localfailures",
  "serverfailures",
  "syncissues",
  "recipientcache",
  "personmetadata",
  "peoplecentricconversationbuddies",
  "imcontactlist",
  "quickcontacts",
  "yammerroot",
  "yammerfeeds",
  "yammerinbound",
  "yammeroutbound",
  "teamchat",
]);

function kindOf(folderClass: string | undefined): FolderKind | "other" {
  if (!folderClass || folderClass === "IPF.Note" || folderClass.startsWith("IPF.Note.")) return "mail";
  if (folderClass.startsWith("IPF.Appointment")) return "calendar";
  if (folderClass.startsWith("IPF.Contact")) return "contacts";
  if (folderClass.startsWith("IPF.Task")) return "tasks";
  return "other";
}

export class FolderService {
  private cache?: { at: number; folders: FolderInfo[] };

  constructor(private readonly client: OwaClient) {}

  /** All folders under the mailbox root, with slash-separated paths. Cached for five minutes. */
  async all(refresh = false): Promise<FolderInfo[]> {
    if (!refresh && this.cache && Date.now() - this.cache.at < CACHE_TTL_MS) return this.cache.folders;
    const [message] = await this.client.callForMessages("FindFolder", {
      FolderShape: { __type: t("FolderResponseShape"), BaseShape: "IdOnly", AdditionalProperties: FOLDER_PROPS.map(propertyUri) },
      ParentFolderIds: [folderIdRef("msgfolderroot")],
      Traversal: "Deep",
      Paging: { __type: t("IndexedPageView"), BasePoint: "Beginning", Offset: 0, MaxEntriesReturned: 1000 },
    });
    const raw: Json[] = (message?.RootFolder as Json | undefined)?.Folders as Json[] ?? [];
    const byId = new Map<string, Json>();
    for (const f of raw) byId.set((f.FolderId as Json).Id as string, f);

    const pathOf = (f: Json, depth = 0): string => {
      const name = String(f.DisplayName ?? "");
      const parent = byId.get(((f.ParentFolderId as Json | undefined)?.Id as string) ?? "");
      return parent && depth < 32 ? `${pathOf(parent, depth + 1)}/${name}` : name;
    };

    const folders = raw.map((f) => normalizeFolder(f, pathOf(f)) as FolderInfo);
    folders.sort((a, b) => a.path.localeCompare(b.path));
    this.cache = { at: Date.now(), folders };
    return folders;
  }

  invalidate(): void {
    this.cache = undefined;
  }

  async list(kind: FolderKind = "mail", refresh = false): Promise<FolderInfo[]> {
    const folders = await this.all(refresh);
    return folders.filter((f) => {
      if (f.wellKnownName && HIDDEN_WELL_KNOWN.has(f.wellKnownName)) return false;
      return kind === "all" || kindOf(f.folderClass) === kind;
    });
  }

  /**
   * Resolve a folder reference to something usable in a request:
   * a well-known name (inbox, sentitems, "sent", "trash", ...), an Exchange folder ID,
   * a display name, or a path such as "Inbox/Projects".
   */
  async resolve(ref: string): Promise<string> {
    const text = ref.trim();
    if (!text) throw new OutlookError("Folder must not be empty");
    if (distinguishedFolder(text)) return distinguishedFolder(text)!;
    if (looksLikeExchangeId(text)) return text;

    const wanted = text.replace(/^\/+|\/+$/g, "").toLowerCase();
    const folders = await this.all();
    const byPath = folders.filter((f) => f.path.toLowerCase() === wanted);
    if (byPath.length === 1) return byPath[0]!.id;
    const byName = folders.filter((f) => (f.name ?? "").toLowerCase() === wanted);
    if (byName.length === 1) return byName[0]!.id;
    const matches = byPath.length ? byPath : byName;
    if (matches.length > 1) {
      throw new OutlookError(
        `Folder "${ref}" is ambiguous; it matches ${matches.map((f) => `"${f.path}"`).join(", ")}. Use the full path or the folder ID.`,
      );
    }
    throw new OutlookError(`Folder "${ref}" was not found. Use list_folders to see available folders.`);
  }

  async create(name: string, parent = "msgfolderroot", folderClass = "IPF.Note"): Promise<FolderInfo> {
    const parentId = await this.resolve(parent);
    const [message] = await this.client.callForMessages("CreateFolder", {
      ParentFolderId: targetFolder(parentId),
      Folders: [{ __type: t("Folder"), FolderClass: folderClass, DisplayName: name }],
    });
    this.invalidate();
    const id = ((message?.Folders as Json[] | undefined)?.[0]?.FolderId as Json | undefined)?.Id as string | undefined;
    if (!id) throw new OutlookError("CreateFolder succeeded but returned no folder ID");
    return (await this.all()).find((f) => f.id === id) ?? { id, name, path: name };
  }

  async rename(ref: string, name: string): Promise<FolderInfo> {
    const id = await this.resolve(ref);
    await this.client.callForMessages("UpdateFolder", {
      FolderChanges: [
        {
          __type: t("FolderChange"),
          FolderId: folderIdRef(id),
          Updates: [{ __type: t("SetFolderField"), Path: propertyUri("folder:DisplayName"), Folder: { __type: t("Folder"), DisplayName: name } }],
        },
      ],
    });
    this.invalidate();
    const folders = await this.all();
    return folders.find((f) => f.id === id || f.wellKnownName === id) ?? { id, name, path: name };
  }

  async move(ref: string, destination: string): Promise<void> {
    const id = await this.resolve(ref);
    const target = await this.resolve(destination);
    await this.client.callForMessages("MoveFolder", { ToFolderId: targetFolder(target), FolderIds: [folderIdRef(id)] });
    this.invalidate();
  }

  async delete(ref: string, permanent: boolean): Promise<void> {
    const id = await this.resolve(ref);
    if (distinguishedFolder(id)) throw new OutlookError(`"${ref}" is a built-in folder and cannot be deleted`);
    await this.client.callForMessages("DeleteFolder", {
      FolderIds: [folderIdRef(id)],
      DeleteType: permanent ? "HardDelete" : "MoveToDeletedItems",
    });
    this.invalidate();
  }
}
