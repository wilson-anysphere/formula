import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { requireOrgMfaSatisfied } from "../auth/mfa";
import { enforceOrgIpAllowlistFromParams } from "../auth/orgIpAllowlist";
import { getClientIp, getUserAgent } from "../http/request-meta";
import {
  assertOutboundRegionAllowed,
  resolvePrimaryStorageRegion,
  DataResidencyViolationError
} from "../policies/dataResidency";
import { isOrgAdmin, type OrgRole } from "../rbac/roles";
import { deleteSecret, putSecret, type SecretStoreKeyring } from "../secrets/secretStore";
import type { MaybeEncryptedSecret, SiemAuthConfig, SiemEndpointConfig } from "../siem/types";
import { requireAuth } from "./auth";

const SECRET_MASK = "***";

const HEADER_NAME_RE = /^[!#$%&'*+\-.^_`|~0-9A-Za-z]+$/;

function clampInt(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.min(max, Math.max(min, Math.trunc(value)));
}

function parseJsonValue(value: unknown): any {
  if (!value) return null;
  if (typeof value === "object") return value;
  if (typeof value === "string") {
    try {
      return JSON.parse(value);
    } catch {
      return null;
    }
  }
  return null;
}

function collectSecretRefsFromConfig(config: SiemEndpointConfig | null): string[] {
  if (!config?.auth) return [];
  const auth = config.auth;
  const refs: string[] = [];
  const maybeAdd = (value: MaybeEncryptedSecret | undefined) => {
    if (!value || typeof value === "string") return;
    if ("secretRef" in value && typeof value.secretRef === "string") refs.push(value.secretRef);
  };

  if (auth.type === "bearer") maybeAdd(auth.token);
  if (auth.type === "basic") {
    maybeAdd(auth.username);
    maybeAdd(auth.password);
  }
  if (auth.type === "header") maybeAdd(auth.value);
  return refs;
}

function maskSecrets(config: SiemEndpointConfig): SiemEndpointConfig {
  if (!config.auth) return config;
  const auth = config.auth;

  if (auth.type === "none") return config;

  if (auth.type === "bearer") {
    return { ...config, auth: { type: "bearer", token: SECRET_MASK } };
  }

  if (auth.type === "basic") {
    return {
      ...config,
      auth: {
        type: "basic",
        username: SECRET_MASK,
        password: SECRET_MASK
      }
    };
  }

  if (auth.type === "header") {
    return {
      ...config,
      auth: {
        type: "header",
        name: auth.name,
        value: SECRET_MASK
      }
    };
  }

  return config;
}

function siemSecretName(orgId: string, kind: "bearerToken" | "basicUsername" | "basicPassword"): string {
  return `siem:${orgId}:${kind}`;
}

function siemHeaderSecretName(orgId: string, headerName: string): string {
  return `siem:${orgId}:headerValue:${headerName.trim().toLowerCase()}`;
}

function normalizeHeaderName(value: string): string {
  return value.trim();
}

function validateHeaderName(value: string, field: string): string {
  const normalized = normalizeHeaderName(value);
  if (normalized.length === 0 || normalized.length > 128) {
    throw new Error(`${field} must be 1-128 chars`);
  }
  if (!HEADER_NAME_RE.test(normalized)) {
    throw new Error(`${field} must be a valid HTTP header token`);
  }
  return normalized;
}

function validateEndpointUrl(raw: string, { requireHttps = false }: { requireHttps?: boolean } = {}): string {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new Error("endpointUrl must be a valid URL");
  }

  if ((process.env.NODE_ENV === "production" || requireHttps) && url.protocol !== "https:") {
    throw new Error("endpointUrl must use https in production");
  }

  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new Error("endpointUrl must use http or https");
  }

  return url.toString();
}

async function requireOrgAdmin(
  request: FastifyRequest,
  reply: FastifyReply,
  orgId: string
): Promise<{ role: OrgRole } | null> {
  if (request.authOrgId && request.authOrgId !== orgId) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  const membership = await request.server.db.query(
    "SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2",
    [orgId, request.user!.id]
  );
  if (membership.rowCount !== 1) {
    reply.code(404).send({ error: "org_not_found" });
    return null;
  }
  const role = membership.rows[0].role as OrgRole;
  if (!isOrgAdmin(role)) {
    reply.code(403).send({ error: "forbidden" });
    return null;
  }
  return { role };
}

const SiemAuthSchema = z.discriminatedUnion("type", [
  z.object({ type: z.literal("none") }),
  z.object({ type: z.literal("bearer"), token: z.string().min(1) }),
  z.object({ type: z.literal("basic"), username: z.string().min(1), password: z.string().min(1) }),
  z.object({ type: z.literal("header"), name: z.string().min(1), value: z.string().min(1) })
]);

type IncomingSiemAuthConfig = z.infer<typeof SiemAuthSchema>;

const SiemRetrySchema = z
  .object({
    maxAttempts: z.number().int().positive().optional(),
    baseDelayMs: z.number().int().positive().optional(),
    maxDelayMs: z.number().int().positive().optional(),
    jitter: z.boolean().optional()
  })
  .strict()
  .optional();

const SiemRedactionSchema = z
  .object({
    redactionText: z.string().min(1).max(100).optional()
  })
  .strict()
  .optional();

const SiemEndpointSchema = z
  .object({
    endpointUrl: z.string().min(1),
    dataRegion: z.string().min(1).optional(),
    format: z.enum(["json", "cef", "leef"]).optional(),
    timeoutMs: z.number().int().positive().optional(),
    idempotencyKeyHeader: z.string().min(1).max(128).nullable().optional(),
    headers: z.record(z.string(), z.string()).optional(),
    auth: SiemAuthSchema.optional(),
    retry: SiemRetrySchema,
    redactionOptions: SiemRedactionSchema,
    batchSize: z.number().int().positive().optional()
  })
  .strict();

const PutBody = z.union([
  z
    .object({
      enabled: z.boolean().optional(),
      config: SiemEndpointSchema
    })
    .strict(),
  // Backwards-compatible: allow the config object itself as the request body.
  SiemEndpointSchema
]);

type IncomingPutBody = z.infer<typeof PutBody>;

async function storeAuth(options: {
  db: FastifyInstance["db"];
  keyring: SecretStoreKeyring;
  orgId: string;
  enabled: boolean;
  auth: IncomingSiemAuthConfig | undefined;
  existingAuth: SiemAuthConfig | undefined;
}): Promise<SiemAuthConfig | undefined> {
  const { db, keyring, orgId, enabled, auth, existingAuth } = options;
  if (!auth) return undefined;

  if (auth.type === "none") return auth;

  if (auth.type === "bearer") {
    const secretName = siemSecretName(orgId, "bearerToken");
    const token = auth.token;

    if (token === SECRET_MASK) {
      if (
        existingAuth?.type === "bearer" &&
        typeof existingAuth.token === "object" &&
        existingAuth.token != null &&
        "secretRef" in existingAuth.token &&
        typeof (existingAuth.token as { secretRef?: unknown }).secretRef === "string"
      ) {
        return existingAuth;
      }
      throw new Error("auth.token is required for bearer auth");
    }

    if (enabled) {
      await putSecret(db, keyring, secretName, token);
    }

    return { type: "bearer", token: { secretRef: secretName } };
  }

  if (auth.type === "basic") {
    const usernameName = siemSecretName(orgId, "basicUsername");
    const passwordName = siemSecretName(orgId, "basicPassword");
    const username = auth.username;
    const password = auth.password;

    let usernameRef: MaybeEncryptedSecret;
    let passwordRef: MaybeEncryptedSecret;

    if (username === SECRET_MASK) {
      if (
        existingAuth?.type === "basic" &&
        typeof existingAuth.username === "object" &&
        existingAuth.username != null &&
        "secretRef" in existingAuth.username &&
        typeof (existingAuth.username as { secretRef?: unknown }).secretRef === "string"
      ) {
        usernameRef = existingAuth.username;
      } else {
        throw new Error("auth.username is required for basic auth");
      }
    } else {
      if (enabled) {
        await putSecret(db, keyring, usernameName, username);
      }
      usernameRef = { secretRef: usernameName };
    }

    if (password === SECRET_MASK) {
      if (
        existingAuth?.type === "basic" &&
        typeof existingAuth.password === "object" &&
        existingAuth.password != null &&
        "secretRef" in existingAuth.password &&
        typeof (existingAuth.password as { secretRef?: unknown }).secretRef === "string"
      ) {
        passwordRef = existingAuth.password;
      } else {
        throw new Error("auth.password is required for basic auth");
      }
    } else {
      if (enabled) {
        await putSecret(db, keyring, passwordName, password);
      }
      passwordRef = { secretRef: passwordName };
    }

    return { type: "basic", username: usernameRef, password: passwordRef };
  }

  if (auth.type === "header") {
    const headerName = validateHeaderName(auth.name, "auth.name");
    const secretName = siemHeaderSecretName(orgId, headerName);

    if (auth.value === SECRET_MASK) {
      const existing = existingAuth;
      if (
        existing?.type === "header" &&
        existing.name.trim().toLowerCase() === headerName.trim().toLowerCase() &&
        typeof existing.value === "object" &&
        existing.value != null &&
        "secretRef" in existing.value &&
        typeof (existing.value as { secretRef?: unknown }).secretRef === "string"
      ) {
        return { type: "header", name: headerName, value: existing.value };
      }

      throw new Error("auth.value is required for header auth");
    }

    if (enabled) {
      await putSecret(db, keyring, secretName, auth.value);
    }

    return { type: "header", name: headerName, value: { secretRef: secretName } };
  }

  return auth;
}

export function registerSiemRoutes(app: FastifyInstance): void {
  app.get(
    "/orgs/:orgId/siem",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const res = await app.db.query("SELECT enabled, config FROM org_siem_configs WHERE org_id = $1", [orgId]);
    if (res.rowCount !== 1) return reply.code(404).send({ error: "siem_config_not_found" });

    const row = res.rows[0] as { enabled: boolean; config: unknown };
    const config = parseJsonValue(row.config) as SiemEndpointConfig | null;
    if (!config || typeof config.endpointUrl !== "string") {
      return reply.code(500).send({ error: "siem_config_invalid" });
    }

    return reply.send({ enabled: Boolean(row.enabled), config: maskSecrets(config) });
    }
  );

  app.put(
    "/orgs/:orgId/siem",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const parsedBody = PutBody.safeParse(request.body);
    if (!parsedBody.success) return reply.code(400).send({ error: "invalid_request" });

    const body = parsedBody.data as IncomingPutBody;
    const incomingConfig = "config" in body ? body.config : body;
    const requestedEnabled = "config" in body ? body.enabled : undefined;

    const existingRes = await app.db.query("SELECT enabled, config FROM org_siem_configs WHERE org_id = $1", [orgId]);
    const existingRow =
      existingRes.rowCount === 1 ? (existingRes.rows[0] as { enabled: boolean; config: unknown }) : null;
    const existingConfig = existingRow ? (parseJsonValue(existingRow.config) as SiemEndpointConfig | null) : null;
    const prevEnabled = existingRow?.enabled ?? null;

    const enabled = requestedEnabled ?? prevEnabled ?? true;

    const settingsRes = await app.db.query<{
      certificate_pinning_enabled: boolean;
      data_residency_region: string;
      data_residency_allowed_regions: unknown;
      allow_cross_region_processing: boolean;
    }>(
      `
        SELECT
          certificate_pinning_enabled,
          data_residency_region,
          data_residency_allowed_regions,
          allow_cross_region_processing
        FROM org_settings
        WHERE org_id = $1
      `,
      [orgId]
    );
    const settingsRow = settingsRes.rowCount === 1 ? (settingsRes.rows[0] as any) : null;
    const certificatePinningEnabled = settingsRow ? Boolean(settingsRow.certificate_pinning_enabled) : false;
    const dataResidencyRegion = settingsRow ? String(settingsRow.data_residency_region ?? "us") : "us";
    const dataResidencyAllowedRegions = settingsRow ? settingsRow.data_residency_allowed_regions : null;
    const allowCrossRegionProcessing = settingsRow ? Boolean(settingsRow.allow_cross_region_processing) : true;

    let effectiveDataRegion: string;
    try {
      const primaryRegion = resolvePrimaryStorageRegion({
        region: dataResidencyRegion,
        allowedRegions: dataResidencyAllowedRegions
      });
      effectiveDataRegion = incomingConfig.dataRegion ?? primaryRegion;

      assertOutboundRegionAllowed({
        orgId,
        requestedRegion: effectiveDataRegion,
        operation: "siem.config.upsert",
        region: dataResidencyRegion,
        allowedRegions: dataResidencyAllowedRegions,
        allowCrossRegionProcessing
      });
    } catch (err) {
      if (err instanceof DataResidencyViolationError) {
        app.metrics.dataResidencyBlockedTotal.inc({ operation: err.operation });
        try {
          await writeAuditEvent(
            app.db,
            createAuditEvent({
              eventType: "org.data_residency.blocked",
              actor: { type: "user", id: request.user!.id },
              context: {
                orgId,
                userId: request.user!.id,
                userEmail: request.user!.email,
                sessionId: request.session?.id ?? null,
                ipAddress: getClientIp(request),
                userAgent: getUserAgent(request)
              },
              resource: { type: "integration", id: "siem", name: "siem" },
              success: false,
              error: { code: "data_residency_violation", message: err.message },
              details: {
                operation: err.operation,
                requestedRegion: err.requestedRegion,
                allowedRegions: err.allowedRegions,
                dataResidencyRegion,
                allowCrossRegionProcessing
              }
            })
          );
        } catch (auditErr) {
          app.log.warn({ err: auditErr, orgId }, "data_residency_blocked_audit_failed");
        }
        return reply.code(400).send({ error: "invalid_request", message: err.message });
      }
      throw err;
    }

    let endpointUrl: string;
    try {
      endpointUrl = validateEndpointUrl(incomingConfig.endpointUrl, {
        // Pinning is only meaningful for https endpoints.
        requireHttps: Boolean(enabled) && certificatePinningEnabled
      });
    } catch {
      return reply.code(400).send({ error: "invalid_request" });
    }

    const config: SiemEndpointConfig = {
      ...incomingConfig,
      endpointUrl,
      dataRegion: effectiveDataRegion,
      batchSize:
        typeof incomingConfig.batchSize === "number" ? clampInt(incomingConfig.batchSize, 1, 1000) : undefined,
      timeoutMs:
        typeof incomingConfig.timeoutMs === "number" ? clampInt(incomingConfig.timeoutMs, 100, 120_000) : undefined
    };

    try {
      if (config.idempotencyKeyHeader) {
        config.idempotencyKeyHeader = validateHeaderName(config.idempotencyKeyHeader, "idempotencyKeyHeader");
      }

      if (config.auth?.type === "header") {
        config.auth = {
          ...config.auth,
          name: validateHeaderName(config.auth.name, "auth.name")
        };
      }

      if (config.headers) {
        const normalized: Record<string, string> = {};
        for (const [key, value] of Object.entries(config.headers)) {
          const header = validateHeaderName(key, "headers");
          normalized[header] = value;
        }
        config.headers = normalized;
      }
    } catch {
      return reply.code(400).send({ error: "invalid_request" });
    }

    const isEnabling = Boolean(enabled) && !Boolean(prevEnabled);
    if (isEnabling) {
      const incomingAuth = incomingConfig.auth;
      if (incomingAuth?.type === "bearer" && incomingAuth.token === SECRET_MASK) {
        return reply.code(400).send({ error: "invalid_request" });
      }
      if (
        incomingAuth?.type === "basic" &&
        (incomingAuth.username === SECRET_MASK || incomingAuth.password === SECRET_MASK)
      ) {
        return reply.code(400).send({ error: "invalid_request" });
      }
      if (incomingAuth?.type === "header" && incomingAuth.value === SECRET_MASK) {
        return reply.code(400).send({ error: "invalid_request" });
      }
    }

    let storedAuth: SiemAuthConfig | undefined;
    try {
      storedAuth = await storeAuth({
        db: app.db,
        keyring: app.config.secretStoreKeys,
        orgId,
        enabled,
        auth: incomingConfig.auth,
        existingAuth: existingConfig?.auth
      });
    } catch {
      return reply.code(400).send({ error: "invalid_request" });
    }

    const storedConfig: SiemEndpointConfig = {
      ...config,
      auth: storedAuth
    };

    const oldRefs = new Set(collectSecretRefsFromConfig(existingConfig));
    const newRefs = new Set(collectSecretRefsFromConfig(storedConfig));

    await app.db.query(
      `
        INSERT INTO org_siem_configs (org_id, enabled, config)
        VALUES ($1, $2, $3)
        ON CONFLICT (org_id)
        DO UPDATE SET enabled = EXCLUDED.enabled, config = EXCLUDED.config, updated_at = now()
      `,
      [orgId, enabled, JSON.stringify(storedConfig)]
    );

    if (!enabled) {
      // Disabling SIEM should remove secrets from the store.
      for (const name of oldRefs) await deleteSecret(app.db, name);
      for (const name of newRefs) await deleteSecret(app.db, name);
    } else {
      for (const name of oldRefs) {
        if (!newRefs.has(name)) await deleteSecret(app.db, name);
      }
    }

    const isAdd = Boolean(enabled) && !Boolean(prevEnabled);
    const isRemove = !Boolean(enabled) && Boolean(prevEnabled);
    const eventType = isAdd
      ? "admin.integration_added"
      : isRemove
        ? "admin.integration_removed"
        : "admin.integration_updated";

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType,
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "integration", id: "siem", name: "siem" },
        success: true,
        details: {
          integration: "siem",
          enabled,
          endpointUrl: config.endpointUrl,
          format: config.format ?? "json",
          dataRegion: storedConfig.dataRegion ?? null
        }
      })
    );

    return reply.send({ enabled, config: maskSecrets(storedConfig) });
    }
  );

  app.delete(
    "/orgs/:orgId/siem",
    { preHandler: [requireAuth, enforceOrgIpAllowlistFromParams] },
    async (request, reply) => {
    const orgId = (request.params as { orgId: string }).orgId;
    const member = await requireOrgAdmin(request, reply, orgId);
    if (!member) return;
    if (request.session && !(await requireOrgMfaSatisfied(app.db, orgId, request.user!))) {
      return reply.code(403).send({ error: "mfa_required" });
    }

    const existingRes = await app.db.query("SELECT enabled, config FROM org_siem_configs WHERE org_id = $1", [orgId]);
    if (existingRes.rowCount !== 1) return reply.code(404).send({ error: "siem_config_not_found" });

    const existingRow = existingRes.rows[0] as { enabled: boolean; config: unknown };
    const existingConfig = parseJsonValue(existingRow.config) as SiemEndpointConfig | null;
    const oldRefs = new Set(collectSecretRefsFromConfig(existingConfig));

    await app.db.query("DELETE FROM org_siem_configs WHERE org_id = $1", [orgId]);
    for (const name of oldRefs) await deleteSecret(app.db, name);

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "admin.integration_removed",
        actor: { type: "user", id: request.user!.id },
        context: {
          orgId,
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id ?? null,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "integration", id: "siem", name: "siem" },
        success: true,
        details: { integration: "siem" }
      })
    );

    return reply.code(204).send();
    }
  );
}
