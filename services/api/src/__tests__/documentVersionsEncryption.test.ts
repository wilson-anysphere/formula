import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { decryptEnvelope, ENVELOPE_VERSION, type EncryptedEnvelope } from "../crypto/envelope";
import { Keyring } from "../crypto/keyring";
import { runKeyRotation } from "../crypto/rotation";
import { createDocumentVersion, getDocumentVersionData } from "../db/documentVersions";
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
  {
    cloudEncryptionAtRest,
    kmsKeyId,
    keyRotationDays,
    kmsKeyRotatedAt
  }: {
    cloudEncryptionAtRest: boolean;
    kmsKeyId?: string | null;
    keyRotationDays?: number;
    kmsKeyRotatedAt?: Date;
  }
): Promise<{ userId: string; orgId: string; docId: string }> {
  const userId = crypto.randomUUID();
  const orgId = crypto.randomUUID();
  const docId = crypto.randomUUID();

  await db.query("INSERT INTO users (id, email, name) VALUES ($1, $2, $3)", [userId, `${userId}@example.com`, "User"]);
  await db.query("INSERT INTO organizations (id, name) VALUES ($1, $2)", [orgId, "Org"]);
  await db.query(
    `
      INSERT INTO org_settings (org_id, cloud_encryption_at_rest, kms_provider, kms_key_id, key_rotation_days, kms_key_rotated_at)
      VALUES ($1, $2, 'local', $3, $4, $5)
    `,
    [
      orgId,
      cloudEncryptionAtRest,
      kmsKeyId ?? null,
      keyRotationDays ?? 90,
      kmsKeyRotatedAt ?? new Date("2026-01-01T00:00:00.000Z")
    ]
  );
  await db.query("INSERT INTO documents (id, org_id, title, created_by) VALUES ($1, $2, $3, $4)", [
    docId,
    orgId,
    "Doc",
    userId
  ]);

  return { userId, orgId, docId };
}

describe("Cloud encryption-at-rest (DB): document_versions.data envelope encryption", () => {
  it("encrypt→store→load→decrypt roundtrip", async () => {
    const db = await setupDb();
    try {
      const keyring = new Keyring({ localMasterKey: "test-master-key", awsKmsEnabled: false, awsRegion: null });
      const { userId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: true });

      const plaintext = Buffer.from("classified", "utf8");
      const created = await createDocumentVersion(db, keyring, {
        documentId: docId,
        createdBy: userId,
        data: plaintext
      });

      const raw = await db.query(
        "SELECT data, data_ciphertext, data_encrypted_dek, data_kms_key_id FROM document_versions WHERE id = $1",
        [created.id]
      );
      expect(raw.rowCount).toBe(1);
      expect(raw.rows[0].data).toBeNull();
      expect(raw.rows[0].data_ciphertext).toBeTypeOf("string");
      expect(raw.rows[0].data_encrypted_dek).toBeTypeOf("string");
      expect(raw.rows[0].data_kms_key_id).toBeTypeOf("string");

      const roundTripped = await getDocumentVersionData(db, keyring, created.id);
      expect(roundTripped?.toString("utf8")).toBe("classified");
    } finally {
      await db.end();
    }
  });

  it("detects AAD mismatch (wrong documentId)", async () => {
    const db = await setupDb();
    try {
      const keyring = new Keyring({ localMasterKey: "test-master-key", awsKmsEnabled: false, awsRegion: null });
      const { userId, orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: true });

      const plaintext = Buffer.from("top-secret", "utf8");
      const created = await createDocumentVersion(db, keyring, {
        documentId: docId,
        createdBy: userId,
        data: plaintext
      });

      const row = (
        await db.query(
          `
            SELECT data_envelope_version, data_algorithm, data_ciphertext, data_iv, data_tag,
                   data_encrypted_dek, data_kms_provider, data_kms_key_id
            FROM document_versions
            WHERE id = $1
          `,
          [created.id]
        )
      ).rows[0] as any;

      const envelope: EncryptedEnvelope = {
        envelopeVersion: Number(row.data_envelope_version),
        algorithm: String(row.data_algorithm),
        ciphertext: Buffer.from(String(row.data_ciphertext), "base64"),
        iv: Buffer.from(String(row.data_iv), "base64"),
        tag: Buffer.from(String(row.data_tag), "base64"),
        encryptedDek: Buffer.from(String(row.data_encrypted_dek), "base64"),
        kmsProvider: String(row.data_kms_provider),
        kmsKeyId: String(row.data_kms_key_id),
        aad: {}
      } as EncryptedEnvelope;

      const correctAad = {
        envelopeVersion: ENVELOPE_VERSION,
        blob: "document_versions.data",
        orgId,
        documentId: docId,
        documentVersionId: created.id
      };
      const wrongAad = { ...correctAad, documentId: crypto.randomUUID() };

      const kms = keyring.get(envelope.kmsProvider);
      const ok = await decryptEnvelope({ envelope, kmsProvider: kms, orgId, aadContext: correctAad });
      expect(ok.toString("utf8")).toBe("top-secret");

      await expect(decryptEnvelope({ envelope, kmsProvider: kms, orgId, aadContext: wrongAad })).rejects.toThrow();
    } finally {
      await db.end();
    }
  });

  it("rotates keys by re-wrapping DEKs (ciphertext unchanged)", async () => {
    const db = await setupDb();
    try {
      const keyring = new Keyring({ localMasterKey: "test-master-key", awsKmsEnabled: false, awsRegion: null });

      const oldKmsKeyId = "local-test-key";
      const rotatedAt = new Date("2026-01-01T00:00:00.000Z");
      const now = new Date("2026-02-10T00:00:00.000Z");

      const { userId, orgId, docId } = await seedOrgAndDoc(db, {
        cloudEncryptionAtRest: true,
        kmsKeyId: oldKmsKeyId,
        keyRotationDays: 1,
        kmsKeyRotatedAt: rotatedAt
      });

      const plaintext = Buffer.from("rotate-me", "utf8");
      const created = await createDocumentVersion(db, keyring, {
        documentId: docId,
        createdBy: userId,
        data: plaintext
      });

      const before = (
        await db.query(
          `
            SELECT data_ciphertext, data_iv, data_tag, data_encrypted_dek, data_kms_key_id
            FROM document_versions
            WHERE id = $1
          `,
          [created.id]
        )
      ).rows[0] as any;

      expect(before.data_kms_key_id).toBe(oldKmsKeyId);

      const rotation = await runKeyRotation(db, keyring, { now });
      expect(rotation.orgsRotated).toBe(1);
      expect(rotation.documentVersionDeksRewrapped).toBe(1);

      const after = (
        await db.query(
          `
            SELECT data_ciphertext, data_iv, data_tag, data_encrypted_dek, data_kms_key_id
            FROM document_versions
            WHERE id = $1
          `,
          [created.id]
        )
      ).rows[0] as any;

      expect(after.data_ciphertext).toBe(before.data_ciphertext);
      expect(after.data_iv).toBe(before.data_iv);
      expect(after.data_tag).toBe(before.data_tag);

      expect(after.data_kms_key_id).not.toBe(oldKmsKeyId);
      expect(after.data_encrypted_dek).not.toBe(before.data_encrypted_dek);

      const roundTripped = await getDocumentVersionData(db, keyring, created.id);
      expect(roundTripped?.toString("utf8")).toBe("rotate-me");

      const orgSettings = await db.query("SELECT kms_key_id, kms_key_rotated_at FROM org_settings WHERE org_id = $1", [
        orgId
      ]);
      expect(orgSettings.rows[0].kms_key_id).toBe(after.data_kms_key_id);
    } finally {
      await db.end();
    }
  });

  it("supports mixed plaintext/encrypted rows when cloud_encryption_at_rest is toggled", async () => {
    const db = await setupDb();
    try {
      const keyring = new Keyring({ localMasterKey: "test-master-key", awsKmsEnabled: false, awsRegion: null });
      const { userId, orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: false });

      const plaintext1 = Buffer.from("plaintext-version", "utf8");
      const v1 = await createDocumentVersion(db, keyring, {
        documentId: docId,
        createdBy: userId,
        data: plaintext1
      });

      const raw1 = await db.query("SELECT data, data_ciphertext FROM document_versions WHERE id = $1", [v1.id]);
      expect(raw1.rows[0].data).toBeInstanceOf(Buffer);
      expect(raw1.rows[0].data_ciphertext).toBeNull();

      await db.query("UPDATE org_settings SET cloud_encryption_at_rest = true WHERE org_id = $1", [orgId]);

      const plaintext2 = Buffer.from("encrypted-version", "utf8");
      const v2 = await createDocumentVersion(db, keyring, {
        documentId: docId,
        createdBy: userId,
        data: plaintext2
      });

      const raw2 = await db.query("SELECT data, data_ciphertext FROM document_versions WHERE id = $1", [v2.id]);
      expect(raw2.rows[0].data).toBeNull();
      expect(raw2.rows[0].data_ciphertext).toBeTypeOf("string");

      expect((await getDocumentVersionData(db, keyring, v1.id))?.toString("utf8")).toBe("plaintext-version");
      expect((await getDocumentVersionData(db, keyring, v2.id))?.toString("utf8")).toBe("encrypted-version");

      // Turning encryption back off must not break reads of already-encrypted rows.
      await db.query("UPDATE org_settings SET cloud_encryption_at_rest = false WHERE org_id = $1", [orgId]);
      expect((await getDocumentVersionData(db, keyring, v2.id))?.toString("utf8")).toBe("encrypted-version");
    } finally {
      await db.end();
    }
  });
});
