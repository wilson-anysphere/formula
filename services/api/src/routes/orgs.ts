import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { enforceOrgIpAllowlistFromParams } from "../auth/orgIpAllowlist";
import { validateDlpPolicy } from "../dlp/dlp";
import { getClientIp, getUserAgent } from "../http/request-meta";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { requireAuth } from "./auth";

function normalizeFingerprintHex(value: string): string {
  return value.replaceAll(":", "").toLowerCase();
}

function isSha256FingerprintHex(value: string): boolean {
  const normalized = normalizeFingerprintHex(value);
  return /^[0-9a-f]{64}$/.test(normalized);
}

function uniqStrings(values: string[] | null | undefined): string[] {
  if (!Array.isArray(values)) return [];
  return Array.from(new Set(values.filter((v) => typeof v === "string" && v.length > 0)));
}

type OrgPolicySnapshot = {
  encryption: {
    cloudEncryptionAtRest: boolean;
    kmsProvider: string;
    kmsKeyId: string | null;
    keyRotationDays: number;
    certificatePinningEnabled: boolean;
    certificatePins: string[];
  };
  dataResidency: {
    region: string;
    allowedRegions: string[];
    allowCrossRegionProcessing: boolean;
    aiProcessingRegion: string | null;
  };
  retention: {
    auditLogRetentionDays: number;
    documentVersionRetentionDays: number;
    deletedDocumentRetentionDays: number;
    legalHoldOverridesRetention: boolean;
  };
};

function extractPolicy(settings: Record<string, any>): OrgPolicySnapshot {
  const residencyRegion = String(settings.data_residency_region ?? "us");
  return {
    encryption: {
      cloudEncryptionAtRest: Boolean(settings.cloud_encryption_at_rest),
      kmsProvider: String(settings.kms_provider ?? "local"),
      kmsKeyId: settings.kms_key_id == null ? null : String(settings.kms_key_id),
      keyRotationDays: Number(settings.key_rotation_days ?? 90),
      certificatePinningEnabled: Boolean(settings.certificate_pinning_enabled),
      certificatePins: uniqStrings(settings.certificate_pins)
    },
    dataResidency: {
      region: residencyRegion,
      allowedRegions:
        residencyRegion === "custom" ? uniqStrings(settings.data_residency_allowed_regions) : [],
      allowCrossRegionProcessing: Boolean(settings.allow_cross_region_processing),
      aiProcessingRegion: settings.ai_processing_region == null ? null : String(settings.ai_processing_region)
    },
    retention: {
      auditLogRetentionDays: Number(settings.audit_log_retention_days ?? 365),
      documentVersionRetentionDays: Number(settings.document_version_retention_days ?? 365),
      deletedDocumentRetentionDays: Number(settings.deleted_document_retention_days ?? 30),
      legalHoldOverridesRetention: Boolean(settings.legal_hold_overrides_retention)
    }
  };
}

function validatePolicy(policy: OrgPolicySnapshot): void {
  const kmsProvider = policy.encryption.kmsProvider;
  if (!["local", "aws", "gcp", "azure"].includes(kmsProvider)) {
    throw new Error(`Unsupported kmsProvider: ${kmsProvider}`);
  }
  if (kmsProvider !== "local" && !policy.encryption.kmsKeyId) {
    throw new Error("kmsKeyId is required when kmsProvider is not local");
  }
  if (!Number.isInteger(policy.encryption.keyRotationDays) || policy.encryption.keyRotationDays <= 0) {
    throw new Error("keyRotationDays must be a positive integer");
  }

  const pins = policy.encryption.certificatePins.map(normalizeFingerprintHex);
  if (policy.encryption.certificatePinningEnabled) {
    if (pins.length === 0) {
      throw new Error("certificatePins must be non-empty when certificate pinning is enabled");
    }
    for (const pin of pins) {
      if (!isSha256FingerprintHex(pin)) {
        throw new Error("certificatePins must be SHA-256 fingerprints (hex, optionally colon-separated)");
      }
    }
  } else {
    for (const pin of pins) {
      if (pin.length > 0 && !isSha256FingerprintHex(pin)) {
        throw new Error("certificatePins must be SHA-256 fingerprints (hex, optionally colon-separated)");
      }
    }
  }

  const region = policy.dataResidency.region;
  const allowedRegions =
    region === "custom"
      ? uniqStrings(policy.dataResidency.allowedRegions)
      : region
        ? [region]
        : [];
  if (allowedRegions.length === 0) {
    throw new Error("data residency requires at least one allowed region");
  }
  if (region !== "custom" && !["us", "eu", "apac"].includes(region)) {
    throw new Error(`Unsupported data residency region: ${region}`);
  }
  if (region === "custom" && allowedRegions.length === 0) {
    throw new Error("custom data residency requires dataResidencyAllowedRegions");
  }

  const aiRegion = policy.dataResidency.aiProcessingRegion;
  if (!policy.dataResidency.allowCrossRegionProcessing && aiRegion && !allowedRegions.includes(aiRegion)) {
    throw new Error(
      `aiProcessingRegion ${aiRegion} violates allowCrossRegionProcessing=false (allowed: ${allowedRegions.join(", ")})`
    );
  }
}

async function requireOrgMember(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<{ role: OrgRole } | null> {
  const membership = await request.server.db.query(
    "SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2",
    [orgId, request.user!.id]
  );
  if (membership.rowCount !== 1) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  return { role: membership.rows[0].role as OrgRole };
}

export function registerOrgRoutes(app: FastifyInstance): void {
  app.get("/orgs", { preHandler: requireAuth }, async (request) => {
    const result = await app.db.query(
      `
        SELECT o.id, o.name, om.role
        FROM organizations o
        JOIN org_members om ON om.org_id = o.id
        WHERE om.user_id = $1
        ORDER BY o.created_at ASC
      `,
      [request.user!.id]
    );
    return {
      organizations: result.rows.map((row) => ({
        id: row.id as string,
        name: row.name as string,
        role: row.role as string
      }))
    };
  });

  app.get("/orgs/:orgId", { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] }, async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgMember(request, reply, orgId);
    if (!member) return;

    const org = await app.db.query("SELECT id, name FROM organizations WHERE id = $1", [orgId]);
    const settings = await app.db.query("SELECT * FROM org_settings WHERE org_id = $1", [orgId]);

    return {
      organization: org.rows[0],
      role: member.role,
      settings: settings.rows[0]
    };
  });

  const PatchSettingsBody = z.object({
    requireMfa: z.boolean().optional(),
    allowedAuthMethods: z.array(z.string()).optional(),
    ipAllowlist: z.array(z.string()).optional(),
    allowExternalSharing: z.boolean().optional(),
    allowPublicLinks: z.boolean().optional(),
    defaultPermission: z.enum(["viewer", "commenter", "editor"]).optional(),
    aiEnabled: z.boolean().optional(),
    aiDataProcessingConsent: z.boolean().optional(),
    dataResidencyRegion: z.string().min(1).optional(),
    dataResidencyAllowedRegions: z.array(z.string().min(1)).optional(),
    allowCrossRegionProcessing: z.boolean().optional(),
    aiProcessingRegion: z.string().min(1).nullable().optional(),
    auditLogRetentionDays: z.number().int().positive().optional(),
    documentVersionRetentionDays: z.number().int().positive().optional(),
    deletedDocumentRetentionDays: z.number().int().positive().optional(),
    legalHoldOverridesRetention: z.boolean().optional(),
    cloudEncryptionAtRest: z.boolean().optional(),
    kmsProvider: z.enum(["local", "aws", "gcp", "azure"]).optional(),
    kmsKeyId: z.string().min(1).nullable().optional(),
    keyRotationDays: z.number().int().positive().optional(),
    certificatePinningEnabled: z.boolean().optional(),
    certificatePins: z.array(z.string().min(1)).optional()
  });

  app.patch(
    "/orgs/:orgId/settings",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const member = await requireOrgMember(request, reply, orgId);
      if (!member) return;
      if (!isOrgAdmin(member.role)) return reply.code(403).send({ error: "forbidden" });

      const parsed = PatchSettingsBody.safeParse(request.body);
      if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

      const updates = parsed.data;

      const currentSettingsRes = await app.db.query("SELECT * FROM org_settings WHERE org_id = $1", [orgId]);
      if (currentSettingsRes.rowCount !== 1) {
        return reply.code(404).send({ error: "org_not_found" });
      }
      const currentSettings = currentSettingsRes.rows[0] as Record<string, any>;
      const beforePolicy = extractPolicy(currentSettings);
      const nextPolicy: OrgPolicySnapshot = structuredClone(beforePolicy);

      if (updates.cloudEncryptionAtRest !== undefined)
        nextPolicy.encryption.cloudEncryptionAtRest = updates.cloudEncryptionAtRest;
      if (updates.kmsProvider !== undefined) nextPolicy.encryption.kmsProvider = updates.kmsProvider;
      if (updates.kmsKeyId !== undefined) nextPolicy.encryption.kmsKeyId = updates.kmsKeyId;
      if (updates.keyRotationDays !== undefined) nextPolicy.encryption.keyRotationDays = updates.keyRotationDays;
      if (updates.certificatePinningEnabled !== undefined)
        nextPolicy.encryption.certificatePinningEnabled = updates.certificatePinningEnabled;
      if (updates.certificatePins !== undefined) nextPolicy.encryption.certificatePins = updates.certificatePins;

      if (updates.dataResidencyRegion !== undefined) nextPolicy.dataResidency.region = updates.dataResidencyRegion;
      if (updates.dataResidencyAllowedRegions !== undefined)
        nextPolicy.dataResidency.allowedRegions = updates.dataResidencyAllowedRegions;
      if (updates.allowCrossRegionProcessing !== undefined)
        nextPolicy.dataResidency.allowCrossRegionProcessing = updates.allowCrossRegionProcessing;
      if (updates.aiProcessingRegion !== undefined)
        nextPolicy.dataResidency.aiProcessingRegion = updates.aiProcessingRegion;

      if (updates.auditLogRetentionDays !== undefined)
        nextPolicy.retention.auditLogRetentionDays = updates.auditLogRetentionDays;
      if (updates.documentVersionRetentionDays !== undefined)
        nextPolicy.retention.documentVersionRetentionDays = updates.documentVersionRetentionDays;
      if (updates.deletedDocumentRetentionDays !== undefined)
        nextPolicy.retention.deletedDocumentRetentionDays = updates.deletedDocumentRetentionDays;
      if (updates.legalHoldOverridesRetention !== undefined)
        nextPolicy.retention.legalHoldOverridesRetention = updates.legalHoldOverridesRetention;

      if (nextPolicy.dataResidency.region !== "custom") {
        nextPolicy.dataResidency.allowedRegions = [];
      }

      try {
        validatePolicy(nextPolicy);
      } catch (err) {
        return reply.code(400).send({ error: "invalid_request", message: (err as Error).message });
      }

      const sets: string[] = [];
      const values: unknown[] = [];
      const addSet = (sql: string, value: unknown) => {
        values.push(value);
        sets.push(`${sql} = $${values.length}`);
      };

      if (updates.requireMfa !== undefined) addSet("require_mfa", updates.requireMfa);
      if (updates.allowedAuthMethods !== undefined)
        addSet("allowed_auth_methods", JSON.stringify(updates.allowedAuthMethods));
      if (updates.ipAllowlist !== undefined) addSet("ip_allowlist", JSON.stringify(updates.ipAllowlist));
      if (updates.allowExternalSharing !== undefined) addSet("allow_external_sharing", updates.allowExternalSharing);
      if (updates.allowPublicLinks !== undefined) addSet("allow_public_links", updates.allowPublicLinks);
      if (updates.defaultPermission !== undefined) addSet("default_permission", updates.defaultPermission);
      if (updates.aiEnabled !== undefined) addSet("ai_enabled", updates.aiEnabled);
      if (updates.aiDataProcessingConsent !== undefined)
        addSet("ai_data_processing_consent", updates.aiDataProcessingConsent);
      if (updates.dataResidencyRegion !== undefined) addSet("data_residency_region", updates.dataResidencyRegion);
      if (updates.dataResidencyAllowedRegions !== undefined) {
        const effectiveRegion =
          updates.dataResidencyRegion !== undefined
            ? updates.dataResidencyRegion
            : String(currentSettings.data_residency_region ?? "us");
        if (effectiveRegion !== "custom") {
          return reply.code(400).send({
            error: "invalid_request",
            message: "dataResidencyAllowedRegions is only valid when dataResidencyRegion is custom"
          });
        }
        addSet("data_residency_allowed_regions", JSON.stringify(updates.dataResidencyAllowedRegions));
      }
      if (updates.allowCrossRegionProcessing !== undefined)
        addSet("allow_cross_region_processing", updates.allowCrossRegionProcessing);
      if (updates.aiProcessingRegion !== undefined) addSet("ai_processing_region", updates.aiProcessingRegion);
      if (updates.auditLogRetentionDays !== undefined)
        addSet("audit_log_retention_days", updates.auditLogRetentionDays);
      if (updates.documentVersionRetentionDays !== undefined)
        addSet("document_version_retention_days", updates.documentVersionRetentionDays);
      if (updates.deletedDocumentRetentionDays !== undefined)
        addSet("deleted_document_retention_days", updates.deletedDocumentRetentionDays);
      if (updates.legalHoldOverridesRetention !== undefined)
        addSet("legal_hold_overrides_retention", updates.legalHoldOverridesRetention);

      if (updates.cloudEncryptionAtRest !== undefined)
        addSet("cloud_encryption_at_rest", updates.cloudEncryptionAtRest);
      if (updates.kmsProvider !== undefined) addSet("kms_provider", updates.kmsProvider);
      if (updates.kmsKeyId !== undefined) addSet("kms_key_id", updates.kmsKeyId);
      if (updates.keyRotationDays !== undefined) addSet("key_rotation_days", updates.keyRotationDays);
      if (updates.certificatePinningEnabled !== undefined)
        addSet("certificate_pinning_enabled", updates.certificatePinningEnabled);
      if (updates.certificatePins !== undefined) addSet("certificate_pins", JSON.stringify(updates.certificatePins));

      // Normalize residency_allowed_regions: only meaningful for `custom`. Clear any stale values on region change.
      if (updates.dataResidencyRegion !== undefined && updates.dataResidencyRegion !== "custom") {
        addSet("data_residency_allowed_regions", null);
      }

      if (sets.length === 0) return reply.send({ ok: true });

      values.push(orgId);
      await app.db.query(
        `
          UPDATE org_settings
          SET ${sets.join(", ")}, updated_at = now()
          WHERE org_id = $${values.length}
        `,
        values
      );

      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "admin.settings_changed",
          actor: { type: "user", id: request.user!.id },
          context: {
            orgId,
            userId: request.user!.id,
            userEmail: request.user!.email,
            sessionId: request.session?.id,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "organization", id: orgId },
          success: true,
          details: { updates }
        })
      );

      const policyChanged = {
        encryption: JSON.stringify(beforePolicy.encryption) !== JSON.stringify(nextPolicy.encryption),
        dataResidency: JSON.stringify(beforePolicy.dataResidency) !== JSON.stringify(nextPolicy.dataResidency),
        retention: JSON.stringify(beforePolicy.retention) !== JSON.stringify(nextPolicy.retention)
      };

      const writePolicyEvent = async (
        section: keyof OrgPolicySnapshot,
        before: OrgPolicySnapshot[keyof OrgPolicySnapshot],
        after: OrgPolicySnapshot[keyof OrgPolicySnapshot]
      ) => {
        await writeAuditEvent(
          app.db,
          createAuditEvent({
            eventType: `org.policy.${String(section)}.updated`,
            actor: { type: "user", id: request.user!.id },
            context: {
              orgId,
              userId: request.user!.id,
              userEmail: request.user!.email,
              sessionId: request.session?.id,
              ipAddress: getClientIp(request),
              userAgent: getUserAgent(request)
            },
            resource: { type: "organization", id: orgId },
            success: true,
            details: { before, after }
          })
        );
      };

      if (policyChanged.encryption) {
        await writePolicyEvent("encryption", beforePolicy.encryption, nextPolicy.encryption);
      }
      if (policyChanged.dataResidency) {
        await writePolicyEvent("dataResidency", beforePolicy.dataResidency, nextPolicy.dataResidency);
      }
      if (policyChanged.retention) {
        await writePolicyEvent("retention", beforePolicy.retention, nextPolicy.retention);
      }

      return reply.send({ ok: true });
    }
  );

  app.get(
    "/orgs/:orgId/dlp-policy",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const member = await requireOrgMember(request, reply, orgId);
      if (!member) return;

      const res = await app.db.query("SELECT policy FROM org_dlp_policies WHERE org_id = $1", [orgId]);
      if (res.rowCount !== 1) return reply.code(404).send({ error: "dlp_policy_not_found" });
      return reply.send({ policy: res.rows[0].policy });
    }
  );

  const PutDlpPolicyBody = z.object({ policy: z.unknown() });

  app.put(
    "/orgs/:orgId/dlp-policy",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
      const orgId = (request.params as { orgId: string }).orgId;
      const member = await requireOrgMember(request, reply, orgId);
      if (!member) return;
      if (!isOrgAdmin(member.role)) return reply.code(403).send({ error: "forbidden" });

      const parsed = PutDlpPolicyBody.safeParse(request.body);
      if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

      try {
        validateDlpPolicy(parsed.data.policy);
      } catch (error) {
        const message = error instanceof Error ? error.message : "Invalid policy";
        return reply.code(400).send({ error: "invalid_policy", message });
      }

      await app.db.query(
        `
          INSERT INTO org_dlp_policies (org_id, policy)
          VALUES ($1, $2)
          ON CONFLICT (org_id)
          DO UPDATE SET policy = EXCLUDED.policy, updated_at = now()
        `,
        [orgId, JSON.stringify(parsed.data.policy)]
      );

      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "admin.settings_changed",
          actor: { type: "user", id: request.user!.id },
          context: {
            orgId,
            userId: request.user!.id,
            userEmail: request.user!.email,
            sessionId: request.session?.id,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "organization", id: orgId },
          success: true,
          details: { updates: { dlpPolicy: true } }
        })
      );

      return reply.send({ policy: parsed.data.policy });
    }
  );
}
