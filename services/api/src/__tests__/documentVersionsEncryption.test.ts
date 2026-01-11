import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { decryptEnvelope } from "../../../../packages/security/crypto/envelope.js";
import { KmsProviderFactory, runKmsRotationSweep } from "../crypto/kms";
import { canonicalJson } from "../crypto/utils";
import {
  createDocumentVersion,
  getDocumentVersionData,
  migrateLegacyEncryptedDocumentVersions
} from "../db/documentVersions";
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
      const kmsFactory = new KmsProviderFactory(db);
      const { userId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: true });

      const plaintext = Buffer.from("classified", "utf8");
      const created = await createDocumentVersion(db, kmsFactory, {
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

      const roundTripped = await getDocumentVersionData(db, kmsFactory, created.id, {
        legacyLocalKmsMasterKey: "test-master-key"
      });
      expect(roundTripped?.toString("utf8")).toBe("classified");
    } finally {
      await db.end();
    }
  });

  it("detects AAD mismatch (wrong documentId)", async () => {
    const db = await setupDb();
    try {
      const kmsFactory = new KmsProviderFactory(db);
      const { userId, orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: true });

      const plaintext = Buffer.from("top-secret", "utf8");
      const created = await createDocumentVersion(db, kmsFactory, {
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

      const wrappedDek = JSON.parse(String(row.data_encrypted_dek));
      const encryptedEnvelope = {
        schemaVersion: 1,
        wrappedDek,
        algorithm: String(row.data_algorithm),
        ciphertext: String(row.data_ciphertext),
        iv: String(row.data_iv),
        tag: String(row.data_tag)
      };

      const correctAad = {
        envelopeVersion: 1,
        blob: "document_versions.data",
        orgId,
        documentId: docId,
        documentVersionId: created.id
      };
      const wrongAad = { ...correctAad, documentId: crypto.randomUUID() };

      const kms = await kmsFactory.forOrgProvider(orgId, "local");
      const ok = await decryptEnvelope({ encryptedEnvelope, kmsProvider: kms, encryptionContext: correctAad });
      expect(ok.toString("utf8")).toBe("top-secret");

      await expect(
        decryptEnvelope({ encryptedEnvelope, kmsProvider: kms, encryptionContext: wrongAad })
      ).rejects.toThrow();
    } finally {
      await db.end();
    }
  });

  it("rotates keys by re-wrapping DEKs (ciphertext unchanged)", async () => {
    const db = await setupDb();
    try {
      const kmsFactory = new KmsProviderFactory(db);
      const rotatedAt = new Date("2026-01-01T00:00:00.000Z");
      const now = new Date("2026-02-10T00:00:00.000Z");

      const { userId, orgId, docId } = await seedOrgAndDoc(db, {
        cloudEncryptionAtRest: true,
        keyRotationDays: 1,
        kmsKeyRotatedAt: rotatedAt
      });

      const plaintext = Buffer.from("rotate-me", "utf8");
      const created = await createDocumentVersion(db, kmsFactory, {
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

      expect(before.data_kms_key_id).toBe("1");

      const rotation = await runKmsRotationSweep(db, { now });
      expect(rotation.rotated).toBe(1);
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

      expect(after.data_kms_key_id).toBe("2");
      expect(after.data_encrypted_dek).not.toBe(before.data_encrypted_dek);

      const roundTripped = await getDocumentVersionData(db, kmsFactory, created.id);
      expect(roundTripped?.toString("utf8")).toBe("rotate-me");

      const orgSettings = await db.query("SELECT kms_key_rotated_at FROM org_settings WHERE org_id = $1", [
        orgId
      ]);
      expect(new Date(orgSettings.rows[0].kms_key_rotated_at as string).getTime()).toBe(now.getTime());
    } finally {
      await db.end();
    }
  });

  it("keeps legacy v1 encrypted rows decryptable", async () => {
    const db = await setupDb();
    try {
      const kmsFactory = new KmsProviderFactory(db);
      const { orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: true });

      const versionId = crypto.randomUUID();
      const aad = {
        envelopeVersion: 1,
        blob: "document_versions.data",
        orgId,
        documentId: docId,
        documentVersionId: versionId
      };

      const dek = crypto.randomBytes(32);
      const payloadAad = Buffer.from(canonicalJson(aad), "utf8");
      const payloadIv = crypto.randomBytes(12);
      const payloadCipher = crypto.createCipheriv("aes-256-gcm", dek, payloadIv, { authTagLength: 16 });
      payloadCipher.setAAD(payloadAad);
      const payloadCiphertext = Buffer.concat([
        payloadCipher.update(Buffer.from("legacy-ciphertext", "utf8")),
        payloadCipher.final()
      ]);
      const payloadTag = payloadCipher.getAuthTag();

      const kmsKeyId = "local-legacy-key";
      const masterKey = crypto.createHash("sha256").update("test-master-key", "utf8").digest();
      const salt = Buffer.from(`formula:local-kms:org:${orgId}`, "utf8");
      const info = Buffer.from(`formula:local-kms:kmsKeyId:${kmsKeyId}`, "utf8");
      const kekRaw = crypto.hkdfSync("sha256", masterKey, salt, info, 32);
      const kek = Buffer.isBuffer(kekRaw) ? kekRaw : Buffer.from(kekRaw);

      const wrapContext = { v: 1, purpose: "dek-wrap", orgId, kmsKeyId };
      const wrapAad = Buffer.from(canonicalJson(wrapContext), "utf8");
      const wrapIv = crypto.randomBytes(12);
      const wrapCipher = crypto.createCipheriv("aes-256-gcm", kek, wrapIv, { authTagLength: 16 });
      wrapCipher.setAAD(wrapAad);
      const wrapCiphertext = Buffer.concat([wrapCipher.update(dek), wrapCipher.final()]);
      const wrapTag = wrapCipher.getAuthTag();
      const encryptedDek = Buffer.concat([Buffer.from([1]), wrapIv, wrapTag, wrapCiphertext]);

      await db.query(
        `
          INSERT INTO document_versions (
            id,
            document_id,
            data,
            data_envelope_version,
            data_algorithm,
            data_ciphertext,
            data_iv,
            data_tag,
            data_encrypted_dek,
            data_kms_provider,
            data_kms_key_id,
            data_aad
          )
          VALUES ($1,$2,NULL,$3,$4,$5,$6,$7,$8,$9,$10,$11)
        `,
        [
          versionId,
          docId,
          1,
          "aes-256-gcm",
          payloadCiphertext.toString("base64"),
          payloadIv.toString("base64"),
          payloadTag.toString("base64"),
          encryptedDek.toString("base64"),
          "local",
          kmsKeyId,
          JSON.stringify(aad)
        ]
      );

      const roundTripped = await getDocumentVersionData(db, kmsFactory, versionId, {
        documentId: docId,
        legacyLocalKmsMasterKey: "test-master-key"
      });
      expect(roundTripped?.toString("utf8")).toBe("legacy-ciphertext");
    } finally {
      await db.end();
    }
  });

  it("migrates legacy v1 encrypted rows to envelope schema v2 by re-wrapping the DEK", async () => {
    const db = await setupDb();
    try {
      const kmsFactory = new KmsProviderFactory(db);
      const { orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: true });

      const versionId = crypto.randomUUID();
      const aad = {
        envelopeVersion: 1,
        blob: "document_versions.data",
        orgId,
        documentId: docId,
        documentVersionId: versionId
      };

      const dek = crypto.randomBytes(32);
      const payloadAad = Buffer.from(canonicalJson(aad), "utf8");
      const payloadIv = crypto.randomBytes(12);
      const payloadCipher = crypto.createCipheriv("aes-256-gcm", dek, payloadIv, { authTagLength: 16 });
      payloadCipher.setAAD(payloadAad);
      const payloadCiphertext = Buffer.concat([payloadCipher.update(Buffer.from("legacy-migrate", "utf8")), payloadCipher.final()]);
      const payloadTag = payloadCipher.getAuthTag();

      const kmsKeyId = "local-legacy-key";
      const masterKey = crypto.createHash("sha256").update("test-master-key", "utf8").digest();
      const salt = Buffer.from(`formula:local-kms:org:${orgId}`, "utf8");
      const info = Buffer.from(`formula:local-kms:kmsKeyId:${kmsKeyId}`, "utf8");
      const kekRaw = crypto.hkdfSync("sha256", masterKey, salt, info, 32);
      const kek = Buffer.isBuffer(kekRaw) ? kekRaw : Buffer.from(kekRaw);

      const wrapContext = { v: 1, purpose: "dek-wrap", orgId, kmsKeyId };
      const wrapAad = Buffer.from(canonicalJson(wrapContext), "utf8");
      const wrapIv = crypto.randomBytes(12);
      const wrapCipher = crypto.createCipheriv("aes-256-gcm", kek, wrapIv, { authTagLength: 16 });
      wrapCipher.setAAD(wrapAad);
      const wrapCiphertext = Buffer.concat([wrapCipher.update(dek), wrapCipher.final()]);
      const wrapTag = wrapCipher.getAuthTag();
      const encryptedDek = Buffer.concat([Buffer.from([1]), wrapIv, wrapTag, wrapCiphertext]);

      await db.query(
        `
          INSERT INTO document_versions (
            id,
            document_id,
            data,
            data_envelope_version,
            data_algorithm,
            data_ciphertext,
            data_iv,
            data_tag,
            data_encrypted_dek,
            data_kms_provider,
            data_kms_key_id,
            data_aad
          )
          VALUES ($1,$2,NULL,$3,$4,$5,$6,$7,$8,$9,$10,$11)
        `,
        [
          versionId,
          docId,
          1,
          "aes-256-gcm",
          payloadCiphertext.toString("base64"),
          payloadIv.toString("base64"),
          payloadTag.toString("base64"),
          encryptedDek.toString("base64"),
          "local",
          kmsKeyId,
          JSON.stringify(aad)
        ]
      );

      const before = (
        await db.query(
          `
            SELECT data_ciphertext, data_iv, data_tag, data_envelope_version, data_encrypted_dek
            FROM document_versions
            WHERE id = $1
          `,
          [versionId]
        )
      ).rows[0] as any;
      expect(before.data_envelope_version).toBe(1);

      const migrated = await migrateLegacyEncryptedDocumentVersions(db, kmsFactory, {
        orgId,
        batchSize: 100,
        legacyLocalKmsMasterKey: "test-master-key"
      });
      expect(migrated.versionsMigrated).toBe(1);

      const after = (
        await db.query(
          `
            SELECT data_ciphertext, data_iv, data_tag, data_envelope_version, data_encrypted_dek, data_kms_key_id
            FROM document_versions
            WHERE id = $1
          `,
          [versionId]
        )
      ).rows[0] as any;

      expect(after.data_ciphertext).toBe(before.data_ciphertext);
      expect(after.data_iv).toBe(before.data_iv);
      expect(after.data_tag).toBe(before.data_tag);
      expect(after.data_envelope_version).toBe(2);
      expect(() => JSON.parse(String(after.data_encrypted_dek))).not.toThrow();
      expect(after.data_kms_key_id).toBe("1");

      const roundTripped = await getDocumentVersionData(db, kmsFactory, versionId, { documentId: docId });
      expect(roundTripped?.toString("utf8")).toBe("legacy-migrate");
    } finally {
      await db.end();
    }
  });

  it("supports mixed plaintext/encrypted rows when cloud_encryption_at_rest is toggled", async () => {
    const db = await setupDb();
    try {
      const kmsFactory = new KmsProviderFactory(db);
      const { userId, orgId, docId } = await seedOrgAndDoc(db, { cloudEncryptionAtRest: false });

      const plaintext1 = Buffer.from("plaintext-version", "utf8");
      const v1 = await createDocumentVersion(db, kmsFactory, {
        documentId: docId,
        createdBy: userId,
        data: plaintext1
      });

      const raw1 = await db.query("SELECT data, data_ciphertext FROM document_versions WHERE id = $1", [v1.id]);
      expect(raw1.rows[0].data).toBeInstanceOf(Buffer);
      expect(raw1.rows[0].data_ciphertext).toBeNull();

      await db.query("UPDATE org_settings SET cloud_encryption_at_rest = true WHERE org_id = $1", [orgId]);

      const plaintext2 = Buffer.from("encrypted-version", "utf8");
      const v2 = await createDocumentVersion(db, kmsFactory, {
        documentId: docId,
        createdBy: userId,
        data: plaintext2
      });

      const raw2 = await db.query("SELECT data, data_ciphertext FROM document_versions WHERE id = $1", [v2.id]);
      expect(raw2.rows[0].data).toBeNull();
      expect(raw2.rows[0].data_ciphertext).toBeTypeOf("string");

      expect((await getDocumentVersionData(db, kmsFactory, v1.id))?.toString("utf8")).toBe("plaintext-version");
      expect((await getDocumentVersionData(db, kmsFactory, v2.id))?.toString("utf8")).toBe("encrypted-version");

      // Turning encryption back off must not break reads of already-encrypted rows.
      await db.query("UPDATE org_settings SET cloud_encryption_at_rest = false WHERE org_id = $1", [orgId]);
      expect((await getDocumentVersionData(db, kmsFactory, v2.id))?.toString("utf8")).toBe("encrypted-version");
    } finally {
      await db.end();
    }
  });
});
