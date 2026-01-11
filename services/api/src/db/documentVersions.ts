import crypto from "node:crypto";
import type { Pool, PoolClient } from "pg";
import { decryptEnvelope, ENVELOPE_VERSION, encryptEnvelope, type EncryptedEnvelope } from "../crypto/envelope";
import { Keyring } from "../crypto/keyring";
import { canonicalJson } from "../crypto/utils";
import { withTransaction } from "./tx";

export type CreateDocumentVersionParams = {
  documentId: string;
  createdBy?: string | null;
  description?: string | null;
  data: Buffer;
};

type OrgEncryptionSettings = {
  cloudEncryptionAtRest: boolean;
  kmsProvider: string;
  kmsKeyId: string | null;
};

function documentVersionDataAad({
  orgId,
  documentId,
  documentVersionId
}: {
  orgId: string;
  documentId: string;
  documentVersionId: string;
}): Record<string, unknown> {
  return {
    envelopeVersion: ENVELOPE_VERSION,
    blob: "document_versions.data",
    orgId,
    documentId,
    documentVersionId
  };
}

async function loadOrgEncryptionSettings(client: PoolClient, orgId: string): Promise<OrgEncryptionSettings> {
  const res = await client.query(
    `
      SELECT cloud_encryption_at_rest, kms_provider, kms_key_id
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
    cloudEncryptionAtRest: Boolean(row.cloud_encryption_at_rest),
    kmsProvider: String(row.kms_provider ?? "local"),
    kmsKeyId: row.kms_key_id == null ? null : String(row.kms_key_id)
  };
}

async function ensureOrgKmsKeyId(
  client: PoolClient,
  {
    orgId,
    kmsProvider,
    kmsKeyId
  }: {
    orgId: string;
    kmsProvider: string;
    kmsKeyId: string | null;
  }
): Promise<string> {
  if (kmsKeyId) return kmsKeyId;

  if (kmsProvider !== "local") {
    throw new Error(`kms_key_id is required for kms_provider=${kmsProvider}`);
  }

  const next = `local-${crypto.randomUUID()}`;
  await client.query(
    `
      UPDATE org_settings
      SET kms_key_id = $2,
          kms_key_rotated_at = now(),
          updated_at = now()
      WHERE org_id = $1
    `,
    [orgId, next]
  );
  return next;
}

export async function createDocumentVersion(
  pool: Pool,
  keyring: Keyring,
  params: CreateDocumentVersionParams
): Promise<{ id: string; createdAt: Date }> {
  const versionId = crypto.randomUUID();
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

    const kmsKeyId = await ensureOrgKmsKeyId(client, {
      orgId,
      kmsProvider: settings.kmsProvider,
      kmsKeyId: settings.kmsKeyId
    });

    const kms = keyring.get(settings.kmsProvider);
    const aad = documentVersionDataAad({ orgId, documentId: params.documentId, documentVersionId: versionId });
    const envelope = await encryptEnvelope({
      plaintext: params.data,
      kmsProvider: kms,
      orgId,
      keyId: kmsKeyId,
      aadContext: aad
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
        envelope.envelopeVersion,
        envelope.algorithm,
        envelope.ciphertext.toString("base64"),
        envelope.iv.toString("base64"),
        envelope.tag.toString("base64"),
        envelope.encryptedDek.toString("base64"),
        envelope.kmsProvider,
        envelope.kmsKeyId,
        JSON.stringify(envelope.aad)
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
  keyring: Keyring,
  versionId: string,
  { documentId: expectedDocumentId }: { documentId?: string } = {}
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
  const storedAad = row.data_aad as Record<string, unknown> | null;
  const expectedAad = documentVersionDataAad({ orgId, documentId, documentVersionId: String(row.id) });

  if (storedAad && canonicalJson(storedAad) !== canonicalJson(expectedAad)) {
    throw new Error("aad_mismatch");
  }

  const ciphertext = Buffer.from(ciphertextBase64, "base64");
  const iv = Buffer.from(String(row.data_iv), "base64");
  const tag = Buffer.from(String(row.data_tag), "base64");
  const encryptedDek = Buffer.from(String(row.data_encrypted_dek), "base64");

  const envelope: EncryptedEnvelope = {
    envelopeVersion: Number(row.data_envelope_version),
    algorithm: String(row.data_algorithm),
    ciphertext,
    iv,
    tag,
    encryptedDek,
    kmsProvider: String(row.data_kms_provider),
    kmsKeyId: String(row.data_kms_key_id),
    aad: storedAad ?? expectedAad
  } as EncryptedEnvelope;

  const kms = keyring.get(envelope.kmsProvider);
  return decryptEnvelope({ envelope, kmsProvider: kms, orgId, aadContext: expectedAad });
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
  keyring: Keyring,
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

      const kmsKeyId = await ensureOrgKmsKeyId(client, {
        orgId: orgIdValue,
        kmsProvider: settings.kmsProvider,
        kmsKeyId: settings.kmsKeyId
      });

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

      const kms = keyring.get(settings.kmsProvider);
      let updated = 0;

      for (const version of plaintextVersions.rows as any[]) {
        const versionIdValue = String(version.id);
        const documentIdValue = String(version.document_id);
        const data = version.data as Buffer;

        const aad = documentVersionDataAad({
          orgId: orgIdValue,
          documentId: documentIdValue,
          documentVersionId: versionIdValue
        });

        const envelope = await encryptEnvelope({
          plaintext: data,
          kmsProvider: kms,
          orgId: orgIdValue,
          keyId: kmsKeyId,
          aadContext: aad
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
            envelope.envelopeVersion,
            envelope.algorithm,
            envelope.ciphertext.toString("base64"),
            envelope.iv.toString("base64"),
            envelope.tag.toString("base64"),
            envelope.encryptedDek.toString("base64"),
            envelope.kmsProvider,
            envelope.kmsKeyId,
            JSON.stringify(envelope.aad)
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
