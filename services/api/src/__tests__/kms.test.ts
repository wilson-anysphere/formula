import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { runMigrations } from "../db/migrations";
import { KmsProviderFactory, rotateOrgKmsKey, runKmsRotationSweep } from "../crypto/kms";
import { decryptEnvelope, encryptEnvelope } from "../../../../packages/security/crypto/envelope.js";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

describe("KMS (integration)", () => {
  async function createDb(): Promise<Pool> {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    const db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });
    return db;
  }

  it("persists local provider state across factory instances", async () => {
    const db = await createDb();
    try {
      const orgId = crypto.randomUUID();
      await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Org"]);
      await db.query("INSERT INTO org_settings (org_id) VALUES ($1)", [orgId]);

      const factory1 = new KmsProviderFactory(db);
      const provider1 = await factory1.forOrg(orgId);

      const dek = crypto.randomBytes(32);
      const wrapped = provider1.wrapKey({ plaintextKey: dek });

      const factory2 = new KmsProviderFactory(db);
      const provider2 = await factory2.forOrg(orgId);
      const unwrapped = provider2.unwrapKey({ wrappedKey: wrapped });
      expect(unwrapped).toEqual(dek);

      const stateRows = await db.query("SELECT org_id FROM org_kms_local_state WHERE org_id = $1", [
        orgId
      ]);
      expect(stateRows.rowCount).toBe(1);
    } finally {
      await db.end();
    }
  });

  it("rotates keys and preserves ability to decrypt older wrapped keys", async () => {
    const db = await createDb();
    try {
      const orgId = crypto.randomUUID();
      await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Org"]);
      await db.query("INSERT INTO org_settings (org_id) VALUES ($1)", [orgId]);

      const factory = new KmsProviderFactory(db);
      const providerV1 = await factory.forOrg(orgId);

      const plaintext = Buffer.from("hello envelope");
      const encryptedV1 = encryptEnvelope({ plaintext, kmsProvider: providerV1 });

      const keyV1 = crypto.randomBytes(32);
      const wrappedV1 = providerV1.wrapKey({ plaintextKey: keyV1 }) as any;
      expect(wrappedV1.kmsKeyVersion).toBe(1);

      await rotateOrgKmsKey(db, orgId);

      const providerV2 = await factory.forOrg(orgId);
      const decryptedV1 = decryptEnvelope({ encryptedEnvelope: encryptedV1, kmsProvider: providerV2 });
      expect(decryptedV1).toEqual(plaintext);

      const keyV2 = crypto.randomBytes(32);
      const wrappedV2 = providerV2.wrapKey({ plaintextKey: keyV2 }) as any;
      expect(wrappedV2.kmsKeyVersion).toBe(2);

      expect(providerV2.unwrapKey({ wrappedKey: wrappedV1 })).toEqual(keyV1);
      expect(providerV2.unwrapKey({ wrappedKey: wrappedV2 })).toEqual(keyV2);

      const audit = await db.query("SELECT event_type FROM audit_log WHERE org_id = $1", [orgId]);
      expect(audit.rows.map((r) => r.event_type)).toEqual(["org.kms.rotated"]);
    } finally {
      await db.end();
    }
  });

  it("rotation sweep rotates orgs only when due", async () => {
    const db = await createDb();
    try {
      const now = new Date("2026-01-10T00:00:00.000Z");
      const old = new Date("2026-01-01T00:00:00.000Z");

      const orgId = crypto.randomUUID();
      await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Org"]);
      await db.query("INSERT INTO org_settings (org_id, key_rotation_days) VALUES ($1, $2)", [orgId, 1]);

      const factory = new KmsProviderFactory(db);
      const providerV1 = await factory.forOrg(orgId);

      const dek = crypto.randomBytes(32);
      const wrappedV1 = providerV1.wrapKey({ plaintextKey: dek });

      // Make the org due for rotation.
      await db.query("UPDATE org_settings SET kms_key_rotated_at = $2 WHERE org_id = $1", [orgId, old]);

      const first = await runKmsRotationSweep(db, { now });
      expect(first).toEqual({ scanned: 1, rotated: 1, failed: 0 });

      const providerAfter = await factory.forOrg(orgId);
      expect(providerAfter.unwrapKey({ wrappedKey: wrappedV1 })).toEqual(dek);

      const second = await runKmsRotationSweep(db, { now });
      expect(second).toEqual({ scanned: 1, rotated: 0, failed: 0 });

      const audit = await db.query("SELECT event_type FROM audit_log WHERE org_id = $1 ORDER BY created_at ASC", [
        orgId
      ]);
      expect(audit.rows.map((r) => r.event_type)).toEqual(["org.kms.rotated"]);
    } finally {
      await db.end();
    }
  });
});
