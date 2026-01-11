import type { Pool } from "pg";
import { getSecret } from "../secrets/secretStore";
import type { MaybeEncryptedSecret, SiemEndpointConfig, SiemAuthConfig } from "./types";

export interface EnabledSiemOrg {
  orgId: string;
  config: SiemEndpointConfig;
}

export interface SiemConfigProvider {
  listEnabledOrgs(): Promise<EnabledSiemOrg[]>;
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

async function resolveSecretValue(
  db: Pool,
  encryptionSecret: string,
  value: MaybeEncryptedSecret | undefined
): Promise<string | undefined> {
  if (!value) return undefined;
  if (typeof value === "string") return value;

  if ("secretRef" in value && typeof value.secretRef === "string") {
    try {
      const resolved = await getSecret(db, encryptionSecret, value.secretRef);
      return resolved ?? undefined;
    } catch {
      return undefined;
    }
  }

  // Backwards-compatible placeholders (pre secret store integration).
  if ("encrypted" in value && typeof value.encrypted === "string") return value.encrypted;
  if ("ciphertext" in value && typeof value.ciphertext === "string") return value.ciphertext;

  return undefined;
}

async function resolveAuthSecrets(
  db: Pool,
  encryptionSecret: string,
  auth: SiemAuthConfig | undefined
): Promise<SiemAuthConfig | undefined | null> {
  if (!auth) return undefined;
  if (auth.type === "none") return auth;

  if (auth.type === "bearer") {
    const token = await resolveSecretValue(db, encryptionSecret, auth.token);
    if (!token) return null;
    return { type: "bearer", token };
  }

  if (auth.type === "basic") {
    const username = await resolveSecretValue(db, encryptionSecret, auth.username);
    const password = await resolveSecretValue(db, encryptionSecret, auth.password);
    if (!username || !password) return null;
    return { type: "basic", username, password };
  }

  if (auth.type === "header") {
    const value = await resolveSecretValue(db, encryptionSecret, auth.value);
    if (!auth.name || !value) return null;
    return { type: "header", name: auth.name, value };
  }

  return null;
}

export class DbSiemConfigProvider implements SiemConfigProvider {
  constructor(
    private readonly db: Pool,
    private readonly encryptionSecret: string,
    private readonly logger: { debug: (...args: any[]) => void; warn: (...args: any[]) => void } = console
  ) {}

  async listEnabledOrgs(): Promise<EnabledSiemOrg[]> {
    try {
      const res = await this.db.query(
        `
          SELECT org_id, config
          FROM org_siem_configs
          WHERE enabled = true
        `
      );

      const enabled: EnabledSiemOrg[] = [];
      for (const row of res.rows) {
        const orgId = String(row.org_id);
        const config = parseConfig(row.config);
        if (!config) continue;

        const resolvedAuth = await resolveAuthSecrets(this.db, this.encryptionSecret, config.auth);
        if (resolvedAuth === null) {
          this.logger.warn({ orgId }, "siem_config_secret_resolution_failed");
          continue;
        }

        enabled.push({
          orgId,
          config: {
            ...config,
            auth: resolvedAuth
          }
        });
      }

      return enabled;
    } catch (err) {
      this.logger.warn({ err }, "siem_config_list_failed");
      return [];
    }
  }
}
