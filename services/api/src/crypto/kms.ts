import path from "node:path";
import { pathToFileURL } from "node:url";
import type { Pool, PoolClient, QueryResult } from "pg";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { withTransaction } from "../db/tx";

// Keep `./crypto/kms` (directory) exports working even though this file shadows it.
export type { DecryptKeyParams, EncryptKeyParams, EncryptKeyResult, KmsProvider } from "./kms/types";
export { AwsKmsProvider } from "./kms/awsKmsProvider";
export { LocalKmsProvider } from "./kms/localKmsProvider";

export type EncryptionContext = Record<string, unknown> | null;

/**
 * KMS provider interface compatible with `packages/security/crypto/envelope.js`.
 */
export interface EnvelopeKmsProvider {
  wrapKey(args: { plaintextKey: Buffer; encryptionContext?: EncryptionContext }): unknown;
  unwrapKey(args: { wrappedKey: unknown; encryptionContext?: EncryptionContext }): Buffer;
}

type DbClient = Pick<Pool, "query">;

type SecurityLocalKmsProviderInstance = EnvelopeKmsProvider & {
  provider: string;
  currentVersion: number;
  rotateKey(): number;
  toJSON(): unknown;
};

type SecurityLocalKmsProviderStatic = {
  new (...args: any[]): SecurityLocalKmsProviderInstance;
  fromJSON(value: unknown): SecurityLocalKmsProviderInstance;
};

const importEsm: (specifier: string) => Promise<any> = new Function(
  "specifier",
  "return import(specifier)"
) as unknown as (specifier: string) => Promise<any>;

let cachedSecurityLocalKmsProvider: Promise<SecurityLocalKmsProviderStatic> | null = null;

async function loadSecurityLocalKmsProvider(): Promise<SecurityLocalKmsProviderStatic> {
  if (cachedSecurityLocalKmsProvider) return cachedSecurityLocalKmsProvider;

  cachedSecurityLocalKmsProvider = (async () => {
    const candidates: string[] = [];
    if (typeof __dirname === "string") {
      candidates.push(
        pathToFileURL(
          path.resolve(__dirname, "../../../../packages/security/crypto/kms/localKmsProvider.js")
        ).href
      );
    }

    candidates.push(
      pathToFileURL(path.resolve(process.cwd(), "packages/security/crypto/kms/localKmsProvider.js")).href,
      pathToFileURL(
        path.resolve(process.cwd(), "..", "..", "packages/security/crypto/kms/localKmsProvider.js")
      ).href
    );

    let lastError: unknown;
    for (const specifier of candidates) {
      try {
        const mod = await importEsm(specifier);
        return mod.LocalKmsProvider as SecurityLocalKmsProviderStatic;
      } catch (err) {
        lastError = err;
      }
    }

    throw lastError instanceof Error ? lastError : new Error("Failed to load LocalKmsProvider");
  })();

  return cachedSecurityLocalKmsProvider;
}

function normalizeJsonValue<T>(value: unknown): T {
  if (typeof value === "string") return JSON.parse(value) as T;
  return value as T;
}

function coerceDate(value: unknown): Date {
  if (value instanceof Date) return value;
  return new Date(String(value));
}

export async function getOrCreateLocalKmsProvider(
  db: DbClient,
  orgId: string,
  { now = new Date() }: { now?: Date } = {}
): Promise<SecurityLocalKmsProviderInstance> {
  const SecurityLocalKmsProvider = await loadSecurityLocalKmsProvider();

  const existing = await db.query<{ provider: unknown }>(
    "SELECT provider FROM org_kms_local_state WHERE org_id = $1",
    [orgId]
  );
  if (existing.rowCount === 1) {
    return SecurityLocalKmsProvider.fromJSON(normalizeJsonValue(existing.rows[0].provider));
  }

  const created = new SecurityLocalKmsProvider();
  const inserted = await db.query(
    `
      INSERT INTO org_kms_local_state (org_id, provider, updated_at)
      VALUES ($1, $2::jsonb, $3)
      ON CONFLICT (org_id) DO NOTHING
    `,
    [orgId, JSON.stringify(created.toJSON()), now]
  );
  if ((inserted.rowCount ?? 0) === 1) return created;

  const after = await db.query<{ provider: unknown }>(
    "SELECT provider FROM org_kms_local_state WHERE org_id = $1",
    [orgId]
  );
  if (after.rowCount !== 1) {
    throw new Error(`Failed to create local KMS state for org ${orgId}`);
  }
  return SecurityLocalKmsProvider.fromJSON(normalizeJsonValue(after.rows[0].provider));
}

export class KmsProviderFactory {
  constructor(private readonly pool: Pool) {}

  async forOrg(orgId: string): Promise<EnvelopeKmsProvider> {
    const settings = await this.pool.query<{ kms_provider: string; kms_key_id: string | null }>(
      "SELECT kms_provider, kms_key_id FROM org_settings WHERE org_id = $1",
      [orgId]
    );
    if (settings.rowCount !== 1) {
      throw new Error(`Missing org_settings row for org ${orgId}`);
    }

    const kmsProvider = String(settings.rows[0].kms_provider ?? "local");
    const kmsKeyId = settings.rows[0].kms_key_id;

    if (kmsProvider === "local") {
      return getOrCreateLocalKmsProvider(this.pool, orgId);
    }

    if (!kmsKeyId) {
      throw new Error(
        `org_settings.kms_key_id is required when kms_provider is ${kmsProvider} (org ${orgId})`
      );
    }

    return new UnimplementedExternalKmsProvider({ kmsProvider, kmsKeyId, orgId });
  }
}

class UnimplementedExternalKmsProvider implements EnvelopeKmsProvider {
  constructor(private readonly options: { kmsProvider: string; kmsKeyId: string; orgId: string }) {}

  wrapKey(): unknown {
    const { kmsProvider, kmsKeyId, orgId } = this.options;
    throw new Error(
      `KMS provider "${kmsProvider}" is configured for org ${orgId} (kms_key_id=${kmsKeyId}), ` +
        "but is not implemented in this reference repo. " +
        "Set org_settings.kms_provider = 'local' for dev/test or implement the provider in services/api/src/crypto/kms.ts."
    );
  }

  unwrapKey(): Buffer {
    const { kmsProvider, kmsKeyId, orgId } = this.options;
    throw new Error(
      `KMS provider "${kmsProvider}" is configured for org ${orgId} (kms_key_id=${kmsKeyId}), ` +
        "but is not implemented in this reference repo. " +
        "Set org_settings.kms_provider = 'local' for dev/test or implement the provider in services/api/src/crypto/kms.ts."
    );
  }
}

type LocalStateRow = { provider: unknown; updated_at: Date };

async function lockLocalStateRow(
  client: PoolClient,
  orgId: string,
  now: Date
): Promise<{ provider: SecurityLocalKmsProviderInstance; updatedAt: Date }> {
  const SecurityLocalKmsProvider = await loadSecurityLocalKmsProvider();

  let res: QueryResult<LocalStateRow> = await client.query(
    "SELECT provider, updated_at FROM org_kms_local_state WHERE org_id = $1 FOR UPDATE",
    [orgId]
  );

  if (res.rowCount !== 1) {
    const created = new SecurityLocalKmsProvider();
    await client.query(
      `
        INSERT INTO org_kms_local_state (org_id, provider, updated_at)
        VALUES ($1, $2::jsonb, $3)
        ON CONFLICT (org_id) DO NOTHING
      `,
      [orgId, JSON.stringify(created.toJSON()), now]
    );

    res = await client.query(
      "SELECT provider, updated_at FROM org_kms_local_state WHERE org_id = $1 FOR UPDATE",
      [orgId]
    );
  }

  if (res.rowCount !== 1) {
    throw new Error(`Failed to lock local KMS state row for org ${orgId}`);
  }

  return {
    provider: SecurityLocalKmsProvider.fromJSON(normalizeJsonValue(res.rows[0].provider)),
    updatedAt: res.rows[0].updated_at
  };
}

async function persistLocalState(
  client: PoolClient,
  orgId: string,
  provider: SecurityLocalKmsProviderInstance,
  now: Date
): Promise<void> {
  await client.query(
    "UPDATE org_kms_local_state SET provider = $2::jsonb, updated_at = $3 WHERE org_id = $1",
    [orgId, JSON.stringify(provider.toJSON()), now]
  );
}

async function lockOrgSettingsRow(
  client: PoolClient,
  orgId: string
): Promise<{ kmsProvider: string; keyRotationDays: number; rotatedAt: Date }> {
  const settings = await client.query<{
    kms_provider: string;
    key_rotation_days: number;
    kms_key_rotated_at: unknown;
  }>("SELECT kms_provider, key_rotation_days, kms_key_rotated_at FROM org_settings WHERE org_id = $1 FOR UPDATE", [
    orgId
  ]);
  if (settings.rowCount !== 1) {
    throw new Error(`Missing org_settings row for org ${orgId}`);
  }

  return {
    kmsProvider: String(settings.rows[0].kms_provider ?? "local"),
    keyRotationDays: Number(settings.rows[0].key_rotation_days ?? 90),
    rotatedAt: coerceDate(settings.rows[0].kms_key_rotated_at)
  };
}

const DAY_MS = 24 * 60 * 60 * 1000;

export async function rotateOrgKmsKey(
  pool: Pool,
  orgId: string,
  { now = new Date(), reason = "manual" }: { now?: Date; reason?: "manual" | "scheduled" } = {}
): Promise<{ provider: string; previousVersion: number; currentVersion: number }> {
  return withTransaction(pool, async (client) => {
    const settings = await lockOrgSettingsRow(client, orgId);
    if (settings.kmsProvider !== "local") {
      throw new Error(
        `rotateOrgKmsKey only supports kms_provider=local in this reference repo (org ${orgId} uses ${settings.kmsProvider})`
      );
    }

    const { provider, updatedAt } = await lockLocalStateRow(client, orgId, now);
    const previousVersion = provider.currentVersion;
    const currentVersion = provider.rotateKey();
    await persistLocalState(client, orgId, provider, now);

    await client.query(
      "UPDATE org_settings SET kms_key_rotated_at = $2, updated_at = now() WHERE org_id = $1",
      [orgId, now]
    );

    await writeAuditEvent(
      client,
      createAuditEvent({
        eventType: "org.kms.rotated",
        timestamp: now.toISOString(),
        actor: { type: "system", id: "kms" },
        context: { orgId },
        resource: { type: "organization", id: orgId },
        success: true,
        details: {
          kmsProvider: provider.provider,
          previousVersion,
          currentVersion,
          reason,
          previousRotationAt: settings.rotatedAt.toISOString(),
          previousStateUpdatedAt: updatedAt.toISOString()
        }
      })
    );

    return { provider: provider.provider, previousVersion, currentVersion };
  });
}

export async function runKmsRotationSweep(
  pool: Pool,
  { now = new Date() }: { now?: Date } = {}
): Promise<{ scanned: number; rotated: number; failed: number }> {
  const orgs = await pool.query<{
    org_id: string;
    key_rotation_days: number;
    kms_key_rotated_at: unknown;
  }>(
    `
      SELECT org_id, key_rotation_days, kms_key_rotated_at
      FROM org_settings
      WHERE kms_provider = 'local'
    `
  );

  let rotated = 0;
  let failed = 0;

  for (const org of orgs.rows) {
    const orgId = String(org.org_id);
    const keyRotationDays = Number(org.key_rotation_days ?? 90);
    if (!Number.isFinite(keyRotationDays) || keyRotationDays <= 0) continue;

    const rotatedAt = coerceDate(org.kms_key_rotated_at);
    const due = now.getTime() - rotatedAt.getTime() >= keyRotationDays * DAY_MS;
    if (!due) continue;

    try {
      const didRotate = await withTransaction(pool, async (client) => {
        const settings = await lockOrgSettingsRow(client, orgId);
        if (settings.kmsProvider !== "local") return false;

        const stillDue =
          now.getTime() - settings.rotatedAt.getTime() >= settings.keyRotationDays * DAY_MS;
        if (!stillDue) return false;

        const { provider } = await lockLocalStateRow(client, orgId, now);
        const previousVersion = provider.currentVersion;
        const currentVersion = provider.rotateKey();
        await persistLocalState(client, orgId, provider, now);

        await client.query(
          "UPDATE org_settings SET kms_key_rotated_at = $2, updated_at = now() WHERE org_id = $1",
          [orgId, now]
        );

        await writeAuditEvent(
          client,
          createAuditEvent({
            eventType: "org.kms.rotated",
            timestamp: now.toISOString(),
            actor: { type: "system", id: "kms" },
            context: { orgId },
            resource: { type: "organization", id: orgId },
            success: true,
            details: {
              kmsProvider: provider.provider,
              previousVersion,
              currentVersion,
              reason: "scheduled",
              previousRotationAt: settings.rotatedAt.toISOString()
            }
          })
        );

        return true;
      });

      if (didRotate) rotated += 1;
    } catch {
      failed += 1;
    }
  }

  return { scanned: orgs.rows.length, rotated, failed };
}

