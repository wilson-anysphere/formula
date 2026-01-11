import type { Pool } from "pg";
import type { SiemEndpointConfig } from "./types";

export interface EnabledSiemOrg {
  orgId: string;
  config: SiemEndpointConfig;
}

export interface SiemConfigProvider {
  listEnabledOrgs(): Promise<EnabledSiemOrg[]>;
}

async function tableExists(db: Pool, tableName: string): Promise<boolean> {
  // `to_regclass` is Postgres-specific and not supported by pg-mem; fall back
  // to information_schema.
  try {
    const res = await db.query("SELECT to_regclass($1) AS reg", [tableName]);
    const reg = res.rows?.[0]?.reg as string | null | undefined;
    return Boolean(reg);
  } catch {
    try {
      const res = await db.query(
        `
          SELECT 1
          FROM information_schema.tables
          WHERE table_schema = 'public' AND table_name = $1
        `,
        [tableName]
      );
      return (res.rowCount ?? 0) > 0;
    } catch {
      return false;
    }
  }
}

function parseConfig(raw: unknown): SiemEndpointConfig | null {
  if (!raw) return null;
  if (typeof raw === "string") {
    try {
      return parseConfig(JSON.parse(raw));
    } catch {
      return null;
    }
  }
  if (!raw || typeof raw !== "object") return null;

  const config = raw as Record<string, unknown>;
  const endpointUrl = config.endpointUrl;
  if (typeof endpointUrl !== "string" || endpointUrl.length === 0) return null;

  return config as unknown as SiemEndpointConfig;
}

/**
 * Best-effort DB-backed config provider.
 *
 * This intentionally tolerates missing schema so the API can run against
 * databases that have not (yet) applied the SIEM config migrations.
 */
export class DbSiemConfigProvider implements SiemConfigProvider {
  private resolvedMode: "unknown" | "org_siem_configs" | "none" = "unknown";

  constructor(
    private readonly db: Pool,
    private readonly logger: { debug: (...args: any[]) => void; warn: (...args: any[]) => void } = console
  ) {}

  async listEnabledOrgs(): Promise<EnabledSiemOrg[]> {
    if (this.resolvedMode === "none") return [];

    if (this.resolvedMode === "unknown") {
      const hasOrgSiemConfigs = await tableExists(this.db, "public.org_siem_configs");
      this.resolvedMode = hasOrgSiemConfigs ? "org_siem_configs" : "none";
    }

    if (this.resolvedMode !== "org_siem_configs") return [];

    try {
      const res = await this.db.query(
        `
          SELECT org_id, config
          FROM org_siem_configs
          WHERE enabled = true
        `
      );

      return res.rows
        .map((row) => {
          const config = parseConfig(row.config);
          if (!config) return null;
          return { orgId: String(row.org_id), config } satisfies EnabledSiemOrg;
        })
        .filter((row): row is EnabledSiemOrg => Boolean(row));
    } catch (err) {
      this.logger.warn({ err }, "siem_config_list_failed");
      return [];
    }
  }
}
