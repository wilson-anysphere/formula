import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { Keyring } from "../crypto/keyring";
import { encryptPlaintextDocumentVersions, getDocumentVersionData } from "../db/documentVersions";
import { runMigrations } from "../db/migrations";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

async function setupDb(): Promise<Pool> {
  const mem = newDb({ autoCreateForeignKeyIndices: true });
  const pgAdapter = mem.adapters.createPg();
  const db = new pgAdapter.Pool();
  await runMigrations(db, { migrationsDir: getMigrationsDir() });
  return db;
}

async function seedOrgAndDoc(
  db: Pool,
  { cloudEncryptionAtRest }: { cloudEncryptionAtRest: boolean }
): Promise<{ userId: string; orgId: string; docId: string }> {
  const userId = crypto.randomUUID();
  const orgId = crypto.randomUUID();
  const docId = crypto.randomUUID();

  await db.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [userId, `${userId}@example.com`, "User"]);
  await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Org"]);
  await db.query("INSERT INTO org_settings (org_id, cloud_encryption_at_rest) VALUES ($1, $2)", [
    orgId,
    cloudEncryptionAtRest
  ]);
  await db.query("INSERT INTO documents (id, org_id, title, created_by) VALUES ($1, $2, $3, $4)", [
    docId,
    orgId,
    "Doc",
    userId
  ]);

  return { userId, orgId, docId };
}

describe("Cloud encryption-at-rest (DB): backfill plaintext document_versions.data", () => {
  it("encrypts existing plaintext rows when cloud_encryption_at_rest is enabled", async () => {
    const db = await setupDb();
    try {
      const keyring = new Keyring({ localMasterKey: "test-master-key", awsKmsEnabled: false, awsRegion: null });
      const { orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: true });

      const versionId = crypto.randomUUID();
      await db.query("INSERT INTO document_versions (id, document_id, data) VALUES ($1, $2, $3)", [
        versionId,
        docId,
        Buffer.from("legacy-plaintext", "utf8")
      ]);

      const result = await encryptPlaintextDocumentVersions(db, keyring, { orgId, batchSize: 10 });
      expect(result).toEqual({ orgsProcessed: 1, versionsEncrypted: 1 });

      const stored = await db.query(
        "SELECT data, data_ciphertext, data_encrypted_dek, data_kms_key_id FROM document_versions WHERE id = $1",
        [versionId]
      );
      expect(stored.rows[0].data).toBeNull();
      expect(stored.rows[0].data_ciphertext).toBeTypeOf("string");
      expect(stored.rows[0].data_encrypted_dek).toBeTypeOf("string");

      const orgSettings = await db.query("SELECT kms_key_id FROM org_settings WHERE org_id = $1", [orgId]);
      expect(orgSettings.rows[0].kms_key_id).toBeTypeOf("string");
      expect(stored.rows[0].data_kms_key_id).toBe(orgSettings.rows[0].kms_key_id);

      const roundTripped = await getDocumentVersionData(db, keyring, versionId, { documentId: docId });
      expect(roundTripped?.toString("utf8")).toBe("legacy-plaintext");

      const again = await encryptPlaintextDocumentVersions(db, keyring, { orgId, batchSize: 10 });
      expect(again).toEqual({ orgsProcessed: 1, versionsEncrypted: 0 });
    } finally {
      await db.end();
    }
  });

  it("skips orgs with cloud_encryption_at_rest disabled", async () => {
    const db = await setupDb();
    try {
      const keyring = new Keyring({ localMasterKey: "test-master-key", awsKmsEnabled: false, awsRegion: null });
      const { orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: false });

      const versionId = crypto.randomUUID();
      await db.query("INSERT INTO document_versions (id, document_id, data) VALUES ($1, $2, $3)", [
        versionId,
        docId,
        Buffer.from("legacy-plaintext", "utf8")
      ]);

      const result = await encryptPlaintextDocumentVersions(db, keyring, { orgId, batchSize: 10 });
      expect(result).toEqual({ orgsProcessed: 0, versionsEncrypted: 0 });

      const stored = await db.query("SELECT data, data_ciphertext FROM document_versions WHERE id = $1", [versionId]);
      expect(stored.rows[0].data).toBeInstanceOf(Buffer);
      expect(stored.rows[0].data_ciphertext).toBeNull();
    } finally {
      await db.end();
    }
  });
});

