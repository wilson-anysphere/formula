import crypto from "node:crypto";
import path from "node:path";
import { pathToFileURL } from "node:url";
import type { Pool, PoolClient } from "pg";
import type { EnvelopeKmsProvider, KmsProviderFactory } from "../crypto/kms";
import { canonicalJson } from "../crypto/utils";
import { withTransaction } from "./tx";

export type CreateDocumentVersionParams = {
  id?: string;
  documentId: string;
  createdBy?: string | null;
  description?: string | null;
  data: Buffer;
};

type OrgEncryptionSettings = {
  cloudEncryptionAtRest: boolean;
};

/**
 * The envelope metadata schema version stored in `document_versions.data_envelope_version`.
 *
 * v1: Legacy services/api envelope encryption + HKDF-based local KMS (kms_key_id-derived).
 * v2: Canonical packages/security envelope encryption + org_kms_local_state local KMS.
 */
const DOCUMENT_VERSION_ENVELOPE_SCHEMA_V1 = 1;
const DOCUMENT_VERSION_ENVELOPE_SCHEMA_V2 = 2;

/**
 * AAD schema version embedded inside the AES-GCM AAD (JSON) for document_versions.data.
 *
 * NOTE: This intentionally remains stable across envelope schema migrations so
 * ciphertext remains valid even if we change the wrapped-DEK metadata format.
 */
const DOCUMENT_VERSION_AAD_VERSION = 1;

type DocumentVersionAad = Record<string, unknown>;

function documentVersionDataAad({
  orgId,
  documentId,
  documentVersionId
}: {
  orgId: string;
  documentId: string;
  documentVersionId: string;
}): DocumentVersionAad {
  return {
    envelopeVersion: DOCUMENT_VERSION_AAD_VERSION,
    blob: "document_versions.data",
    orgId,
    documentId,
    documentVersionId
  };
}

type SecurityEncryptedEnvelope = {
  schemaVersion: number;
  wrappedDek: unknown;
  algorithm: string;
  iv: string;
  ciphertext: string;
  tag: string;
};

type SecurityEnvelopeCrypto = {
  encryptEnvelope(args: {
    plaintext: Buffer;
    kmsProvider: EnvelopeKmsProvider;
    encryptionContext?: DocumentVersionAad | null;
  }): Promise<SecurityEncryptedEnvelope>;
  decryptEnvelope(args: {
    encryptedEnvelope: SecurityEncryptedEnvelope;
    kmsProvider: EnvelopeKmsProvider;
    encryptionContext?: DocumentVersionAad | null;
  }): Promise<Buffer>;
};

const importEsm: (specifier: string) => Promise<any> = new Function(
  "specifier",
  "return import(specifier)"
) as unknown as (specifier: string) => Promise<any>;

let cachedEnvelopeCrypto: Promise<SecurityEnvelopeCrypto> | null = null;

async function loadEnvelopeCrypto(): Promise<SecurityEnvelopeCrypto> {
  if (cachedEnvelopeCrypto) return cachedEnvelopeCrypto;

  cachedEnvelopeCrypto = (async () => {
    const candidates: string[] = [];
    if (typeof __dirname === "string") {
      candidates.push(pathToFileURL(path.resolve(__dirname, "../../../../packages/security/crypto/envelope.js")).href);
    }

    candidates.push(
      pathToFileURL(path.resolve(process.cwd(), "packages/security/crypto/envelope.js")).href,
      pathToFileURL(path.resolve(process.cwd(), "..", "..", "packages/security/crypto/envelope.js")).href
    );

    let lastError: unknown;
    for (const specifier of candidates) {
      try {
        const mod = await importEsm(specifier);
        return { encryptEnvelope: mod.encryptEnvelope, decryptEnvelope: mod.decryptEnvelope } as SecurityEnvelopeCrypto;
      } catch (err) {
        lastError = err;
      }
    }

    throw lastError instanceof Error ? lastError : new Error("Failed to load envelope crypto");
  })();

  return cachedEnvelopeCrypto;
}

function normalizeJsonValue<T>(value: unknown): T {
  if (typeof value === "string") return JSON.parse(value) as T;
  return value as T;
}

const AES_256_KEY_BYTES = 32;
const AES_GCM_IV_BYTES = 12;
const AES_GCM_TAG_BYTES = 16;

function assertBufferLength(buf: Buffer, expected: number, name: string): void {
  if (!Buffer.isBuffer(buf)) {
    throw new TypeError(`${name} must be a Buffer`);
  }
  if (buf.length !== expected) {
    throw new RangeError(`${name} must be ${expected} bytes (got ${buf.length})`);
  }
}

function aadBytes(context: unknown | null | undefined): Buffer | null {
  if (context === null || context === undefined) return null;
  return Buffer.from(canonicalJson(context), "utf8");
}

function decryptAes256Gcm({
  ciphertext,
  key,
  iv,
  tag,
  aad = null
}: {
  ciphertext: Buffer;
  key: Buffer;
  iv: Buffer;
  tag: Buffer;
  aad?: Buffer | null;
}): Buffer {
  assertBufferLength(key, AES_256_KEY_BYTES, "key");
  assertBufferLength(iv, AES_GCM_IV_BYTES, "iv");
  assertBufferLength(tag, AES_GCM_TAG_BYTES, "tag");
  if (aad !== null && !Buffer.isBuffer(aad)) {
    throw new TypeError("aad must be a Buffer when provided");
  }

  const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv, { authTagLength: AES_GCM_TAG_BYTES });
  decipher.setAuthTag(tag);
  if (aad) {
    decipher.setAAD(aad);
  }
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]);
}

function deriveLegacyLocalMasterKey(secret: string): Buffer {
  if (!secret) {
    throw new Error("LOCAL_KMS_MASTER_KEY must be set to decrypt legacy local envelope rows");
  }
  return crypto.createHash("sha256").update(secret, "utf8").digest();
}

type LegacyWrappedDekV1 = { iv: Buffer; tag: Buffer; ciphertext: Buffer };

function decodeLegacyWrappedDekV1(blob: Buffer): LegacyWrappedDekV1 {
  if (!Buffer.isBuffer(blob)) {
    throw new TypeError("encryptedDek must be a Buffer");
  }
  const minLength = 1 + AES_GCM_IV_BYTES + AES_GCM_TAG_BYTES + 1;
  if (blob.length < minLength) {
    throw new RangeError(`encryptedDek too short (got ${blob.length} bytes)`);
  }

  const version = blob.readUInt8(0);
  if (version !== 1) {
    throw new Error(`Unsupported wrapped DEK version: ${version}`);
  }

  const ivStart = 1;
  const tagStart = ivStart + AES_GCM_IV_BYTES;
  const ciphertextStart = tagStart + AES_GCM_TAG_BYTES;
  return {
    iv: blob.subarray(ivStart, tagStart),
    tag: blob.subarray(tagStart, ciphertextStart),
    ciphertext: blob.subarray(ciphertextStart)
  };
}

function deriveLegacyLocalKek({
  masterKey,
  orgId,
  kmsKeyId
}: {
  masterKey: Buffer;
  orgId: string;
  kmsKeyId: string;
}): Buffer {
  assertBufferLength(masterKey, AES_256_KEY_BYTES, "masterKey");
  if (!orgId) throw new Error("orgId is required");
  if (!kmsKeyId) throw new Error("kmsKeyId is required");

  const salt = Buffer.from(`formula:local-kms:org:${orgId}`, "utf8");
  const info = Buffer.from(`formula:local-kms:kmsKeyId:${kmsKeyId}`, "utf8");
  const derived = crypto.hkdfSync("sha256", masterKey, salt, info, AES_256_KEY_BYTES);
  const buf = Buffer.isBuffer(derived) ? derived : Buffer.from(derived);
  assertBufferLength(buf, AES_256_KEY_BYTES, "kek");
  return buf;
}

function legacyDekWrapAad(orgId: string, kmsKeyId: string): Buffer {
  return aadBytes({
    v: 1,
    purpose: "dek-wrap",
    orgId,
    kmsKeyId
  })!;
}

function unwrapLegacyLocalDek({
  encryptedDek,
  orgId,
  kmsKeyId,
  localKmsMasterKey
}: {
  encryptedDek: Buffer;
  orgId: string;
  kmsKeyId: string;
  localKmsMasterKey: string;
}): Buffer {
  const masterKey = deriveLegacyLocalMasterKey(localKmsMasterKey);
  const parsed = decodeLegacyWrappedDekV1(encryptedDek);
  const kek = deriveLegacyLocalKek({ masterKey, orgId, kmsKeyId });
  const plaintextDek = decryptAes256Gcm({
    ciphertext: parsed.ciphertext,
    key: kek,
    iv: parsed.iv,
    tag: parsed.tag,
    aad: legacyDekWrapAad(orgId, kmsKeyId)
  });
  assertBufferLength(plaintextDek, AES_256_KEY_BYTES, "plaintextDek");
  return plaintextDek;
}

async function loadOrgEncryptionSettings(client: PoolClient, orgId: string): Promise<OrgEncryptionSettings> {
  const res = await client.query(
    `
      SELECT cloud_encryption_at_rest
      FROM org_settings
      WHERE org_id = $1
      LIMIT 1
    `,
    [orgId]
  );

  if (res.rowCount !== 1) {
    // Org created outside the API might not have settings yet; insert defaults.
    await client.query("INSERT INTO org_settings (org_id) VALUES ($1) ON CONFLICT (org_id) DO NOTHING", [orgId]);
    return loadOrgEncryptionSettings(client, orgId);
  }

  const row = res.rows[0] as any;
  return {
    cloudEncryptionAtRest: Boolean(row.cloud_encryption_at_rest)
  };
}

export type EncryptedDocumentVersionData = {
  envelopeVersion: typeof DOCUMENT_VERSION_ENVELOPE_SCHEMA_V2;
  algorithm: string;
  ciphertext: string;
  iv: string;
  tag: string;
  encryptedDek: string;
  kmsProvider: string;
  kmsKeyId: string | null;
  aad: DocumentVersionAad;
};

type DbClient = Pick<Pool, "query">;

export async function encryptDocumentVersionData({
  plaintext,
  kmsFactory,
  orgId,
  documentId,
  documentVersionId,
  db
}: {
  plaintext: Buffer;
  kmsFactory: KmsProviderFactory;
  orgId: string;
  documentId: string;
  documentVersionId: string;
  /**
   * Optional DB client used to resolve org_settings + local KMS state.
   *
   * Pass the current transaction client when encrypting inside a transaction to
   * avoid read-after-write issues (e.g. when org_settings is inserted in the
   * same transaction).
   */
  db?: DbClient;
}): Promise<EncryptedDocumentVersionData> {
  const kms = await kmsFactory.forOrg(orgId, db);
  const { encryptEnvelope } = await loadEnvelopeCrypto();
  const aad = documentVersionDataAad({ orgId, documentId, documentVersionId });
  const envelope = await encryptEnvelope({ plaintext, kmsProvider: kms, encryptionContext: aad });

  const wrappedDekAny = envelope.wrappedDek as any;
  const dataKmsProvider = String(wrappedDekAny?.kmsProvider ?? kms.provider);
  const dataKmsKeyId =
    typeof wrappedDekAny?.kmsKeyId === "string"
      ? wrappedDekAny.kmsKeyId
      : typeof wrappedDekAny?.kmsKeyVersion === "number"
        ? String(wrappedDekAny.kmsKeyVersion)
        : null;

  return {
    envelopeVersion: DOCUMENT_VERSION_ENVELOPE_SCHEMA_V2,
    algorithm: envelope.algorithm,
    ciphertext: envelope.ciphertext,
    iv: envelope.iv,
    tag: envelope.tag,
    encryptedDek: JSON.stringify(envelope.wrappedDek),
    kmsProvider: dataKmsProvider,
    kmsKeyId: dataKmsKeyId,
    aad
  };
}

export async function createDocumentVersion(
  pool: Pool,
  kmsFactory: KmsProviderFactory,
  params: CreateDocumentVersionParams
): Promise<{ id: string; createdAt: Date }> {
  const versionId = params.id ?? crypto.randomUUID();
  let createdAt: Date | null = null;

  await withTransaction(pool, async (client) => {
    const docRes = await client.query("SELECT org_id FROM documents WHERE id = $1 LIMIT 1", [params.documentId]);
    if (docRes.rowCount !== 1) {
      throw new Error("document_not_found");
    }
    const orgId = String((docRes.rows[0] as any).org_id);

    const settings = await loadOrgEncryptionSettings(client, orgId);

    if (!settings.cloudEncryptionAtRest) {
      const inserted = await client.query(
        `
          INSERT INTO document_versions (id, document_id, created_by, description, data)
          VALUES ($1, $2, $3, $4, $5)
          RETURNING created_at
        `,
        [versionId, params.documentId, params.createdBy ?? null, params.description ?? null, params.data]
      );
      createdAt = (inserted.rows[0] as any).created_at as Date;
      return;
    }

    const encrypted = await encryptDocumentVersionData({
      plaintext: params.data,
      kmsFactory,
      orgId,
      documentId: params.documentId,
      documentVersionId: versionId,
      db: client
    });

    const inserted = await client.query(
      `
        INSERT INTO document_versions (
          id,
          document_id,
          created_by,
          description,
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
         VALUES (
           $1,$2,$3,$4,
           NULL,
           $5,$6,$7,$8,$9,$10,$11,$12,$13
         )
         RETURNING created_at
       `,
      [
        versionId,
        params.documentId,
        params.createdBy ?? null,
        params.description ?? null,
        encrypted.envelopeVersion,
        encrypted.algorithm,
        encrypted.ciphertext,
        encrypted.iv,
        encrypted.tag,
        encrypted.encryptedDek,
        encrypted.kmsProvider,
        encrypted.kmsKeyId,
        JSON.stringify(encrypted.aad)
      ]
    );
    createdAt = (inserted.rows[0] as any).created_at as Date;
  });

  if (!createdAt) {
    throw new Error("failed_to_create_document_version");
  }

  return { id: versionId, createdAt };
}

export async function getDocumentVersionData(
  pool: Pool,
  kmsFactory: KmsProviderFactory,
  versionId: string,
  {
    documentId: expectedDocumentId,
    legacyLocalKmsMasterKey
  }: { documentId?: string; legacyLocalKmsMasterKey?: string } = {}
): Promise<Buffer | null> {
  const whereDoc = expectedDocumentId ? "AND v.document_id = $2" : "";
  const params = expectedDocumentId ? [versionId, expectedDocumentId] : [versionId];
  const res = await pool.query(
    `
      SELECT
        v.id,
        v.document_id,
        d.org_id,
        v.data,
        v.data_envelope_version,
        v.data_algorithm,
        v.data_ciphertext,
        v.data_iv,
        v.data_tag,
        v.data_encrypted_dek,
        v.data_kms_provider,
        v.data_kms_key_id,
        v.data_aad
      FROM document_versions v
      JOIN documents d ON d.id = v.document_id
      WHERE v.id = $1
      ${whereDoc}
      LIMIT 1
    `,
    params
  );

  if (res.rowCount !== 1) return null;

  const row = res.rows[0] as any;
  const plaintext = row.data as Buffer | null;
  const ciphertextBase64 = row.data_ciphertext as string | null;

  if (!ciphertextBase64) {
    return plaintext;
  }

  const orgId = String(row.org_id);
  const documentId = String(row.document_id);
  const storedAad = row.data_aad == null ? null : normalizeJsonValue<Record<string, unknown>>(row.data_aad);
  const expectedAad = documentVersionDataAad({ orgId, documentId, documentVersionId: String(row.id) });

  if (storedAad && canonicalJson(storedAad) !== canonicalJson(expectedAad)) {
    throw new Error("aad_mismatch");
  }

  const envelopeVersion = row.data_envelope_version == null ? null : Number(row.data_envelope_version);

  if (envelopeVersion === DOCUMENT_VERSION_ENVELOPE_SCHEMA_V2) {
    const wrappedDek = JSON.parse(String(row.data_encrypted_dek));
    const encryptedEnvelope: SecurityEncryptedEnvelope = {
      schemaVersion: 1,
      wrappedDek,
      algorithm: String(row.data_algorithm),
      ciphertext: String(row.data_ciphertext),
      iv: String(row.data_iv),
      tag: String(row.data_tag)
    };

    const providerName = String((wrappedDek as any)?.kmsProvider ?? row.data_kms_provider ?? "local");
    const kms = await kmsFactory.forOrgProvider(orgId, providerName);
    const { decryptEnvelope } = await loadEnvelopeCrypto();
    return decryptEnvelope({ encryptedEnvelope, kmsProvider: kms, encryptionContext: expectedAad });
  }

  // Legacy v1: data_encrypted_dek is base64 bytes; unwrap/decrypt locally.
  if (envelopeVersion !== DOCUMENT_VERSION_ENVELOPE_SCHEMA_V1) {
    throw new Error(`Unsupported document_versions envelope schema version: ${String(envelopeVersion)}`);
  }

  const ciphertext = Buffer.from(String(row.data_ciphertext), "base64");
  const iv = Buffer.from(String(row.data_iv), "base64");
  const tag = Buffer.from(String(row.data_tag), "base64");
  const encryptedDek = Buffer.from(String(row.data_encrypted_dek), "base64");

  const legacyProvider = String(row.data_kms_provider ?? "local");
  const legacyKmsKeyId = String(row.data_kms_key_id ?? "");

  let dek: Buffer;
  if (legacyProvider === "local") {
    dek = unwrapLegacyLocalDek({
      encryptedDek,
      orgId,
      kmsKeyId: legacyKmsKeyId,
      localKmsMasterKey: legacyLocalKmsMasterKey ?? ""
    });
  } else if (legacyProvider === "aws") {
    const kms = await kmsFactory.forOrgProvider(orgId, "aws");
    dek = await kms.unwrapKey({
      wrappedKey: { kmsProvider: "aws", kmsKeyId: legacyKmsKeyId, ciphertext: encryptedDek.toString("base64") },
      encryptionContext: expectedAad
    });
  } else {
    throw new Error(`Unsupported legacy kms provider: ${legacyProvider}`);
  }

  return decryptAes256Gcm({
    ciphertext,
    key: dek,
    iv,
    tag,
    aad: aadBytes(expectedAad)
  });
}

export type EncryptPlaintextDocumentVersionsResult = {
  orgsProcessed: number;
  versionsEncrypted: number;
};

/**
 * Backfill helper to encrypt existing plaintext document_versions.data rows for
 * orgs that have cloud_encryption_at_rest enabled.
 *
 * This is intentionally batch-oriented so it can be run periodically.
 */
export async function encryptPlaintextDocumentVersions(
  pool: Pool,
  kmsFactory: KmsProviderFactory,
  {
    orgId,
    batchSize = 100
  }: {
    orgId?: string;
    batchSize?: number;
  } = {}
): Promise<EncryptPlaintextDocumentVersionsResult> {
  const limit = Number(batchSize ?? 100);
  if (!Number.isInteger(limit) || limit <= 0) {
    throw new Error("batchSize must be a positive integer");
  }

  const orgs = await pool.query(
    `
      SELECT org_id
      FROM org_settings
      ${orgId ? "WHERE org_id = $1" : ""}
      ORDER BY org_id ASC
    `,
    orgId ? [orgId] : []
  );

  let orgsProcessed = 0;
  let versionsEncrypted = 0;

  for (const row of orgs.rows as any[]) {
    const orgIdValue = String(row.org_id);

    const encryptedForOrg = await withTransaction(pool, async (client) => {
      const settings = await loadOrgEncryptionSettings(client, orgIdValue);
      if (!settings.cloudEncryptionAtRest) return 0;

      orgsProcessed += 1;

      const plaintextVersions = await client.query(
        `
          SELECT v.id, v.document_id, v.data
          FROM document_versions v
          JOIN documents d ON d.id = v.document_id
          WHERE d.org_id = $1
            AND v.data IS NOT NULL
            AND v.data_ciphertext IS NULL
          ORDER BY v.created_at ASC
          LIMIT $2
        `,
        [orgIdValue, limit]
      );

      if ((plaintextVersions.rowCount ?? 0) === 0) return 0;

      let updated = 0;

      for (const version of plaintextVersions.rows as any[]) {
        const versionIdValue = String(version.id);
        const documentIdValue = String(version.document_id);
        const data = version.data as Buffer;

        const encrypted = await encryptDocumentVersionData({
          plaintext: data,
          kmsFactory,
          orgId: orgIdValue,
          documentId: documentIdValue,
          documentVersionId: versionIdValue,
          db: client
        });

        await client.query(
          `
            UPDATE document_versions
            SET data = NULL,
                data_envelope_version = $2,
                data_algorithm = $3,
                data_ciphertext = $4,
                data_iv = $5,
                data_tag = $6,
                data_encrypted_dek = $7,
                data_kms_provider = $8,
                data_kms_key_id = $9,
                data_aad = $10
            WHERE id = $1
          `,
          [
            versionIdValue,
            encrypted.envelopeVersion,
            encrypted.algorithm,
            encrypted.ciphertext,
            encrypted.iv,
            encrypted.tag,
            encrypted.encryptedDek,
            encrypted.kmsProvider,
            encrypted.kmsKeyId,
            JSON.stringify(encrypted.aad)
          ]
        );

        updated += 1;
      }

      return updated;
    });

    versionsEncrypted += encryptedForOrg;
  }

  return { orgsProcessed, versionsEncrypted };
}

export type MigrateLegacyEncryptedDocumentVersionsResult = {
  orgsProcessed: number;
  versionsMigrated: number;
};

/**
 * Upgrade legacy envelope schema v1 rows (HKDF local KMS + base64 wrapped DEK)
 * to the canonical envelope schema v2 representation (JSON wrapped-key object).
 *
 * This migration re-wraps the existing DEK under the org's *current* KMS provider
 * and does **not** re-encrypt the ciphertext payload.
 */
export async function migrateLegacyEncryptedDocumentVersions(
  pool: Pool,
  kmsFactory: KmsProviderFactory,
  {
    orgId,
    batchSize = 100,
    legacyLocalKmsMasterKey
  }: {
    orgId?: string;
    batchSize?: number;
    /**
     * Required to migrate legacy v1 rows that were encrypted with the historical
     * HKDF-based local KMS provider.
     */
    legacyLocalKmsMasterKey?: string;
  } = {}
): Promise<MigrateLegacyEncryptedDocumentVersionsResult> {
  const limit = Number(batchSize ?? 100);
  if (!Number.isInteger(limit) || limit <= 0) {
    throw new Error("batchSize must be a positive integer");
  }

  const orgs = await pool.query(
    `
      SELECT org_id
      FROM org_settings
      ${orgId ? "WHERE org_id = $1" : ""}
      ORDER BY org_id ASC
    `,
    orgId ? [orgId] : []
  );

  let orgsProcessed = 0;
  let versionsMigrated = 0;

  for (const row of orgs.rows as any[]) {
    const orgIdValue = String(row.org_id);

    const migratedForOrg = await withTransaction(pool, async (client) => {
      // Ensure org_settings exists (some tests insert org_settings manually).
      await loadOrgEncryptionSettings(client, orgIdValue);

      const legacyVersions = await client.query(
        `
          SELECT v.id, v.document_id, v.data_encrypted_dek, v.data_kms_provider, v.data_kms_key_id, v.data_aad
          FROM document_versions v
          JOIN documents d ON d.id = v.document_id
          WHERE d.org_id = $1
            AND v.data_ciphertext IS NOT NULL
            AND v.data_envelope_version = $2
            AND v.data_encrypted_dek IS NOT NULL
          ORDER BY v.created_at ASC
          LIMIT $3
        `,
        [orgIdValue, DOCUMENT_VERSION_ENVELOPE_SCHEMA_V1, limit]
      );

      if ((legacyVersions.rowCount ?? 0) === 0) return 0;

      orgsProcessed += 1;

      // Target provider for re-wrapping is whatever the org is currently
      // configured to use.
      const targetKms = await kmsFactory.forOrg(orgIdValue, client);

      let updated = 0;

      for (const version of legacyVersions.rows as any[]) {
        const versionIdValue = String(version.id);
        const documentIdValue = String(version.document_id);
        const expectedAad = documentVersionDataAad({
          orgId: orgIdValue,
          documentId: documentIdValue,
          documentVersionId: versionIdValue
        });

        const storedAad =
          version.data_aad == null ? null : normalizeJsonValue<Record<string, unknown>>(version.data_aad);
        if (storedAad && canonicalJson(storedAad) !== canonicalJson(expectedAad)) {
          throw new Error("aad_mismatch");
        }

        const legacyProvider = String(version.data_kms_provider ?? "local");
        const legacyKmsKeyId = String(version.data_kms_key_id ?? "");
        const encryptedDekBytes = Buffer.from(String(version.data_encrypted_dek), "base64");

        let dek: Buffer;
        if (legacyProvider === "local") {
          dek = unwrapLegacyLocalDek({
            encryptedDek: encryptedDekBytes,
            orgId: orgIdValue,
            kmsKeyId: legacyKmsKeyId,
            localKmsMasterKey: legacyLocalKmsMasterKey ?? ""
          });
        } else if (legacyProvider === "aws") {
          const kms = await kmsFactory.forOrgProvider(orgIdValue, "aws", client);
          dek = await kms.unwrapKey({
            wrappedKey: {
              kmsProvider: "aws",
              kmsKeyId: legacyKmsKeyId,
              ciphertext: encryptedDekBytes.toString("base64")
            },
            encryptionContext: expectedAad
          });
        } else {
          throw new Error(`Unsupported legacy kms provider: ${legacyProvider}`);
        }

        const wrappedDek = await targetKms.wrapKey({ plaintextKey: dek, encryptionContext: expectedAad });
        const wrappedDekAny = wrappedDek as any;
        const dataKmsProvider = String(wrappedDekAny?.kmsProvider ?? targetKms.provider);
        const dataKmsKeyId =
          typeof wrappedDekAny?.kmsKeyId === "string"
            ? wrappedDekAny.kmsKeyId
            : typeof wrappedDekAny?.kmsKeyVersion === "number"
              ? String(wrappedDekAny.kmsKeyVersion)
              : null;

        await client.query(
          `
            UPDATE document_versions
            SET data_envelope_version = $2,
                data_encrypted_dek = $3,
                data_kms_provider = $4,
                data_kms_key_id = $5,
                data_aad = COALESCE(data_aad, $6::jsonb)
            WHERE id = $1
          `,
          [
            versionIdValue,
            DOCUMENT_VERSION_ENVELOPE_SCHEMA_V2,
            JSON.stringify(wrappedDek),
            dataKmsProvider,
            dataKmsKeyId,
            JSON.stringify(expectedAad)
          ]
        );

        updated += 1;
      }

      return updated;
    });

    versionsMigrated += migratedForOrg;
  }

  return { orgsProcessed, versionsMigrated };
}
