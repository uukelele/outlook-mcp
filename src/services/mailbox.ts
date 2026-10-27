import { t } from "../owa/builders.js";
import type { OwaClient } from "../owa/client.js";
import { compact } from "../owa/normalize.js";

export interface MailboxInfo {
  displayName?: string;
  email?: string;
  userPrincipalName?: string;
  tenantDomain?: string;
  culture?: string;
  /** Windows time zone ID from the mailbox settings. */
  timeZone?: string;
  maxMessageSizeKb?: number;
  quota?: { usedBytes?: number; warningBytes?: number; sendLimitBytes?: number };
}

export interface Category {
  name: string;
  color: string;
  id?: string;
}

const CACHE_TTL_MS = 15 * 60 * 1000;

export class MailboxService {
  private cache?: { at: number; config: any };
  private inflight?: Promise<any>;

  constructor(
    private readonly client: OwaClient,
    private readonly timeZoneOverride?: string,
  ) {}

  private async configuration(refresh = false): Promise<any> {
    if (!refresh && this.cache && Date.now() - this.cache.at < CACHE_TTL_MS) return this.cache.config;
    this.inflight ??= this.client
      .call(
        "GetOwaUserConfiguration",
        {
          __type: t("GetOwaUserConfigurationJsonRequest"),
          Header: { __type: t("JsonRequestHeaders"), RequestServerVersion: "V2018_01_08" },
          Body: { __type: t("GetOwaUserConfigurationRequest"), Owaconfigs: ["MasterCategoryList"] },
        },
        { raw: true },
      )
      .then((config) => {
        this.cache = { at: Date.now(), config };
        return config;
      })
      .finally(() => {
        this.inflight = undefined;
      });
    return this.inflight;
  }

  async info(): Promise<MailboxInfo> {
    const config = await this.configuration();
    const session = config?.SessionSettings ?? {};
    const number = (value: unknown) => (typeof value === "number" && value >= 0 ? value : undefined);
    return compact({
      displayName: session.UserDisplayName,
      email: session.UserEmailAddress,
      userPrincipalName: session.UserPrincipalName,
      tenantDomain: session.TenantDomain,
      culture: session.UserCulture,
      timeZone: await this.timeZone(),
      maxMessageSizeKb: number(session.MaxMessageSizeInKb),
      quota: {
        usedBytes: number(session.QuotaUsed),
        warningBytes: number(session.QuotaWarning),
        sendLimitBytes: number(session.QuotaSend),
      },
    });
  }

  /** SMTP address of the signed-in mailbox. */
  async emailAddress(): Promise<string | undefined> {
    const session = (await this.configuration())?.SessionSettings ?? {};
    return session.UserEmailAddress ?? session.UserPrincipalName ?? this.client.identity()?.upn;
  }

  /** Windows time zone ID used to interpret wall-clock times (OUTLOOK_TIMEZONE overrides the mailbox setting). */
  async timeZone(): Promise<string> {
    if (this.timeZoneOverride) return this.timeZoneOverride;
    try {
      const zone = (await this.configuration())?.UserOptions?.TimeZone;
      return typeof zone === "string" && zone ? zone : "UTC";
    } catch {
      return "UTC";
    }
  }

  async categories(): Promise<Category[]> {
    const config = await this.configuration(true);
    const list: any[] = config?.MasterCategoryList?.MasterList ?? [];
    return list
      .filter((c) => typeof c?.Name === "string")
      .map((c) => compact({ name: c.Name, color: categoryColor(c.Color), id: c.Id }) as Category)
      .sort((a, b) => a.name.localeCompare(b.name));
  }
}

/** Outlook category colors are preset indexes; -1 means no color. */
function categoryColor(value: unknown): string {
  return typeof value === "number" && value >= 0 ? `preset${value}` : "none";
}
