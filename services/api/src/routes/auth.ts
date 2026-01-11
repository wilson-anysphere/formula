import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { createAuditEvent, writeAuditEvent } from "../audit/audit";
import { authenticateApiKey } from "../auth/apiKeys";
import { generateTotpSecret, buildOtpAuthUrl, verifyTotpCode } from "../auth/mfa";
import { oidcCallback, oidcStart } from "../auth/oidc/oidc";
import { hashPassword, verifyPassword } from "../auth/password";
import { samlCallback, samlStart } from "../auth/saml/saml";
import { createSession, lookupSessionByToken, revokeSession } from "../auth/sessions";
import { withTransaction } from "../db/tx";
import { TokenBucketRateLimiter, sha256Hex } from "../http/rateLimit";
import { getClientIp, getUserAgent } from "../http/request-meta";

type AuthCredentials =
  | { kind: "session"; token: string }
  | { kind: "api_key"; token: string };

function extractAuthCredentials(request: FastifyRequest): AuthCredentials | null {
  const cookieName = request.server.config.sessionCookieName;
  const cookieToken = request.cookies?.[cookieName];
  if (cookieToken && typeof cookieToken === "string") return { kind: "session", token: cookieToken };

  const xApiKey = request.headers["x-api-key"];
  if (typeof xApiKey === "string" && xApiKey.trim().length > 0) {
    const raw = xApiKey.trim();
    return { kind: "api_key", token: raw.startsWith("api_") ? raw : `api_${raw}` };
  }

  const auth = request.headers.authorization;
  if (!auth || typeof auth !== "string") return null;
  const [kind, token] = auth.split(" ");
  if (kind?.toLowerCase() !== "bearer") return null;
  if (!token) return null;
  if (token.startsWith("api_")) return { kind: "api_key", token };
  return { kind: "session", token };
}

async function requireAuth(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const credentials = extractAuthCredentials(request);
  if (!credentials) {
    request.server.metrics.authFailuresTotal.inc({ reason: "missing_token" });
    reply.code(401).send({ error: "unauthorized" });
    return;
  }

  if (credentials.kind === "session") {
    const found = await lookupSessionByToken(request.server.db, credentials.token);
    if (!found) {
      request.server.metrics.authFailuresTotal.inc({ reason: "invalid_token" });
      reply.code(401).send({ error: "unauthorized" });
      return;
    }

    request.user = found.user;
    request.session = found.session;
    request.apiKey = undefined;
    request.authMethod = "session";
    request.authOrgId = undefined;
    return;
  }

  const apiKeyResult = await authenticateApiKey(request.server.db, credentials.token, {
    clientIp: getClientIp(request)
  });

  if (!apiKeyResult.ok) {
    request.server.metrics.authFailuresTotal.inc({ reason: apiKeyResult.value.error });
    reply.code(apiKeyResult.value.statusCode).send({ error: apiKeyResult.value.error });
    return;
  }

  request.user = apiKeyResult.value.user;
  request.session = undefined;
  request.apiKey = apiKeyResult.value.apiKey;
  request.authMethod = "api_key";
  request.authOrgId = apiKeyResult.value.apiKey.orgId;

  await writeAuditEvent(
    request.server.db,
    createAuditEvent({
      eventType: "auth.api_key_used",
      actor: { type: "api_key", id: apiKeyResult.value.apiKey.id },
      context: {
        orgId: apiKeyResult.value.apiKey.orgId,
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      },
      resource: { type: "api_key", id: apiKeyResult.value.apiKey.id, name: apiKeyResult.value.apiKey.name },
      success: true,
      details: { createdBy: apiKeyResult.value.apiKey.createdBy }
    })
  );
}

export function registerAuthRoutes(app: FastifyInstance): void {
  const loginIpLimiter = new TokenBucketRateLimiter({ capacity: 20, refillMs: 60_000 });
  const loginEmailLimiter = new TokenBucketRateLimiter({ capacity: 5, refillMs: 60_000 });
  const registerIpLimiter = new TokenBucketRateLimiter({ capacity: 20, refillMs: 60_000 });
  const oidcIpLimiter = new TokenBucketRateLimiter({ capacity: 30, refillMs: 60_000 });
  const samlIpLimiter = new TokenBucketRateLimiter({ capacity: 30, refillMs: 60_000 });

  const rateLimited = (reply: FastifyReply, retryAfterMs: number) => {
    return reply
      .header("Retry-After", String(Math.max(1, Math.ceil(retryAfterMs / 1000))))
      .code(429)
      .send({ error: "too_many_requests" });
  };

  const RegisterBody = z.object({
    email: z.string().email(),
    password: z.string().min(8),
    name: z.string().min(1),
    orgName: z.string().min(1).optional()
  });

  app.post("/auth/register", async (request, reply) => {
    const body = RegisterBody.safeParse(request.body);
    if (!body.success) return reply.code(400).send({ error: "invalid_request" });

    const email = body.data.email.trim().toLowerCase();
    const name = body.data.name.trim();
    const password = body.data.password;

    const ip = getClientIp(request) ?? "unknown";
    const limited = registerIpLimiter.take(ip);
    if (!limited.ok) {
      app.metrics.rateLimitedTotal.inc({ route: "/auth/register", reason: "ip" });
      return rateLimited(reply, limited.retryAfterMs);
    }

    const existing = await app.db.query("SELECT 1 FROM users WHERE email = $1", [email]);
    if (existing.rowCount && existing.rowCount > 0) {
      return reply.code(409).send({ error: "email_in_use" });
    }

    const userId = crypto.randomUUID();
    const orgId = crypto.randomUUID();
    const orgName = body.data.orgName?.trim() ?? `${name}'s org`;
    const passwordHash = await hashPassword(password);
    const now = new Date();
    const sessionExpiresAt = new Date(now.getTime() + app.config.sessionTtlSeconds * 1000);

    const { sessionId, token } = await withTransaction(app.db, async (client) => {
      await client.query(
        `
          INSERT INTO users (id, email, name, password_hash)
          VALUES ($1, $2, $3, $4)
        `,
        [userId, email, name, passwordHash]
      );

      await client.query(
        `
          INSERT INTO organizations (id, name)
          VALUES ($1, $2)
        `,
        [orgId, orgName]
      );

      await client.query(
        `
          INSERT INTO org_settings (org_id)
          VALUES ($1)
        `,
        [orgId]
      );

      await client.query(
        `
          INSERT INTO org_members (org_id, user_id, role)
          VALUES ($1, $2, 'owner')
        `,
        [orgId, userId]
      );

      return createSession(client, {
        userId,
        expiresAt: sessionExpiresAt,
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      });
    });

    reply.setCookie(app.config.sessionCookieName, token, {
      path: "/",
      httpOnly: true,
      sameSite: "lax",
      secure: app.config.cookieSecure
    });

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "auth.login",
        actor: { type: "user", id: userId },
        context: {
          orgId,
          userId,
          userEmail: email,
          sessionId,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "session", id: sessionId },
        success: true,
        details: { method: "password", operation: "register" }
      })
    );

    return reply.send({
      user: { id: userId, email, name },
      organization: { id: orgId, name: orgName }
    });
  });

  const LoginBody = z.object({
    email: z.string().email(),
    password: z.string().min(1),
    mfaCode: z.string().min(1).optional()
  });

  app.post("/auth/login", async (request, reply) => {
    const body = LoginBody.safeParse(request.body);
    if (!body.success) return reply.code(400).send({ error: "invalid_request" });

    const email = body.data.email.trim().toLowerCase();
    const password = body.data.password;

    const ip = getClientIp(request) ?? "unknown";
    const ipResult = loginIpLimiter.take(ip);
    const emailResult = loginEmailLimiter.take(sha256Hex(email));
    if (!ipResult.ok || !emailResult.ok) {
      const retryAfterMs = Math.max(ipResult.ok ? 0 : ipResult.retryAfterMs, emailResult.ok ? 0 : emailResult.retryAfterMs);
      const reason = !ipResult.ok && !emailResult.ok ? "multiple" : !ipResult.ok ? "ip" : "email";
      app.metrics.rateLimitedTotal.inc({ route: "/auth/login", reason });
      return rateLimited(reply, retryAfterMs);
    }

    const found = await app.db.query(
      `
        SELECT id, email, name, password_hash, mfa_totp_secret, mfa_totp_enabled
        FROM users
        WHERE email = $1
        LIMIT 1
      `,
      [email]
    );

    if (found.rowCount !== 1) {
      app.metrics.authFailuresTotal.inc({ reason: "invalid_credentials" });
      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "auth.login_failed",
          actor: { type: "anonymous", id: email },
          context: {
            userEmail: email,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "user", id: null },
          success: false,
          error: { code: "invalid_credentials" },
          details: { method: "password" }
        })
      );
      return reply.code(401).send({ error: "invalid_credentials" });
    }

    const row = found.rows[0] as {
      id: string;
      email: string;
      name: string;
      password_hash: string | null;
      mfa_totp_secret: string | null;
      mfa_totp_enabled: boolean;
    };

    if (!row.password_hash) {
      app.metrics.authFailuresTotal.inc({ reason: "password_login_disabled" });
      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "auth.login_failed",
          actor: { type: "user", id: row.id },
          context: {
            userId: row.id,
            userEmail: row.email,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "user", id: row.id },
          success: false,
          error: { code: "password_login_disabled" },
          details: { method: "password" }
        })
      );
      return reply.code(401).send({ error: "invalid_credentials" });
    }

    const ok = await verifyPassword(password, row.password_hash);
    if (!ok) {
      app.metrics.authFailuresTotal.inc({ reason: "invalid_credentials" });
      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "auth.login_failed",
          actor: { type: "user", id: row.id },
          context: {
            userId: row.id,
            userEmail: row.email,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "user", id: row.id },
          success: false,
          error: { code: "invalid_credentials" },
          details: { method: "password" }
        })
      );
      return reply.code(401).send({ error: "invalid_credentials" });
    }

    if (row.mfa_totp_enabled) {
      const code = body.data.mfaCode;
      if (!code || !row.mfa_totp_secret || !verifyTotpCode(row.mfa_totp_secret, code)) {
        app.metrics.authFailuresTotal.inc({ reason: "mfa_required" });
        await writeAuditEvent(
          app.db,
          createAuditEvent({
            eventType: "auth.login_failed",
            actor: { type: "user", id: row.id },
            context: {
              userId: row.id,
              userEmail: row.email,
              ipAddress: getClientIp(request),
              userAgent: getUserAgent(request)
            },
            resource: { type: "user", id: row.id },
            success: false,
            error: { code: "mfa_required" },
            details: { method: "password" }
          })
        );
        return reply.code(401).send({ error: "mfa_required" });
      }
    }

    const now = new Date();
    const expiresAt = new Date(now.getTime() + app.config.sessionTtlSeconds * 1000);
    const { sessionId, token } = await withTransaction(app.db, async (client) =>
      createSession(client, {
        userId: row.id,
        expiresAt,
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      })
    );

    reply.setCookie(app.config.sessionCookieName, token, {
      path: "/",
      httpOnly: true,
      sameSite: "lax",
      secure: app.config.cookieSecure
    });

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "auth.login",
        actor: { type: "user", id: row.id },
        context: {
          userId: row.id,
          userEmail: row.email,
          sessionId,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "session", id: sessionId },
        success: true,
        details: { method: "password" }
      })
    );

    return reply.send({ user: { id: row.id, email: row.email, name: row.name } });
  });

  app.post("/auth/logout", { preHandler: requireAuth }, async (request, reply) => {
    const sessionId = request.session?.id;
    if (sessionId) {
      await revokeSession(app.db, sessionId);
      await writeAuditEvent(
        app.db,
        createAuditEvent({
          eventType: "auth.logout",
          actor: { type: "user", id: request.user?.id ?? "unknown" },
          context: {
            userId: request.user?.id ?? null,
            userEmail: request.user?.email ?? null,
            sessionId,
            ipAddress: getClientIp(request),
            userAgent: getUserAgent(request)
          },
          resource: { type: "session", id: sessionId },
          success: true,
          details: {}
        })
      );
    }

    reply.clearCookie(app.config.sessionCookieName, { path: "/" });
    return reply.send({ ok: true });
  });

  app.get("/me", { preHandler: requireAuth }, async (request) => {
    const orgFilter = request.authOrgId ? "AND o.id = $2" : "";
    const params = request.authOrgId ? [request.user!.id, request.authOrgId] : [request.user!.id];
    const orgs = await app.db.query(
      `
        SELECT o.id, o.name, om.role
        FROM organizations o
        JOIN org_members om ON om.org_id = o.id
        WHERE om.user_id = $1
          ${orgFilter}
        ORDER BY o.created_at ASC
      `,
      params
    );

    return {
      user: request.user,
      apiKey: request.apiKey
        ? {
            id: request.apiKey.id,
            orgId: request.apiKey.orgId,
            name: request.apiKey.name
          }
        : null,
      organizations: orgs.rows.map((row) => ({
        id: row.id as string,
        name: row.name as string,
        role: row.role as string
      }))
    };
  });

  // MFA scaffolding (TOTP). Secrets are stored server-side; production deployments should encrypt at rest.
  app.post("/auth/mfa/totp/setup", { preHandler: requireAuth }, async (request) => {
    const secret = generateTotpSecret();
    const otpauthUrl = buildOtpAuthUrl({
      issuer: "Formula",
      accountName: request.user!.email,
      secret
    });

    await app.db.query("UPDATE users SET mfa_totp_secret = $1, mfa_totp_enabled = false WHERE id = $2", [
      secret,
      request.user!.id
    ]);

    return { secret, otpauthUrl };
  });

  const TotpConfirmBody = z.object({ code: z.string().min(1) });

  app.post("/auth/mfa/totp/confirm", { preHandler: requireAuth }, async (request, reply) => {
    const parsed = TotpConfirmBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const secretRow = await app.db.query("SELECT mfa_totp_secret FROM users WHERE id = $1", [request.user!.id]);
    const secret = (secretRow.rows[0]?.mfa_totp_secret ?? null) as string | null;
    if (!secret || !verifyTotpCode(secret, parsed.data.code)) {
      return reply.code(400).send({ error: "invalid_code" });
    }

    await app.db.query("UPDATE users SET mfa_totp_enabled = true WHERE id = $1", [request.user!.id]);
    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "auth.mfa_enabled",
        actor: { type: "user", id: request.user!.id },
        context: {
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "user", id: request.user!.id },
        success: true,
        details: {}
      })
    );

    return reply.send({ ok: true });
  });

  app.post("/auth/mfa/totp/disable", { preHandler: requireAuth }, async (request, reply) => {
    const parsed = TotpConfirmBody.safeParse(request.body);
    if (!parsed.success) return reply.code(400).send({ error: "invalid_request" });

    const secretRow = await app.db.query(
      "SELECT mfa_totp_secret, mfa_totp_enabled FROM users WHERE id = $1",
      [request.user!.id]
    );
    const secret = (secretRow.rows[0]?.mfa_totp_secret ?? null) as string | null;
    const enabled = Boolean(secretRow.rows[0]?.mfa_totp_enabled);

    if (enabled && (!secret || !verifyTotpCode(secret, parsed.data.code))) {
      return reply.code(400).send({ error: "invalid_code" });
    }

    await app.db.query("UPDATE users SET mfa_totp_enabled = false, mfa_totp_secret = null WHERE id = $1", [
      request.user!.id
    ]);

    await writeAuditEvent(
      app.db,
      createAuditEvent({
        eventType: "auth.mfa_disabled",
        actor: { type: "user", id: request.user!.id },
        context: {
          userId: request.user!.id,
          userEmail: request.user!.email,
          sessionId: request.session?.id,
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        },
        resource: { type: "user", id: request.user!.id },
        success: true,
        details: {}
      })
    );

    return reply.send({ ok: true });
  });

  // OIDC / SSO: per-organization providers.
  const oidcRateLimitByIp = (route: string) => async (request: FastifyRequest, reply: FastifyReply) => {
    const ip = getClientIp(request) ?? "unknown";
    const limited = oidcIpLimiter.take(ip);
    if (!limited.ok) {
      app.metrics.rateLimitedTotal.inc({ route, reason: "ip" });
      return rateLimited(reply, limited.retryAfterMs);
    }
  };

  const samlRateLimitByIp = (route: string) => async (request: FastifyRequest, reply: FastifyReply) => {
    const ip = getClientIp(request) ?? "unknown";
    const limited = samlIpLimiter.take(ip);
    if (!limited.ok) {
      app.metrics.rateLimitedTotal.inc({ route, reason: "ip" });
      return rateLimited(reply, limited.retryAfterMs);
    }
  };

  app.get(
    "/auth/oidc/:orgId/:provider/start",
    { preHandler: oidcRateLimitByIp("/auth/oidc/:orgId/:provider/start") },
    oidcStart
  );
  app.get(
    "/auth/oidc/:orgId/:provider/callback",
    { preHandler: oidcRateLimitByIp("/auth/oidc/:orgId/:provider/callback") },
    oidcCallback
  );

  // SAML 2.0 SSO: per-organization providers.
  app.get(
    "/auth/saml/:orgId/:provider/start",
    { preHandler: samlRateLimitByIp("/auth/saml/:orgId/:provider/start") },
    samlStart
  );
  app.post(
    "/auth/saml/:orgId/:provider/callback",
    { preHandler: samlRateLimitByIp("/auth/saml/:orgId/:provider/callback") },
    samlCallback
  );
}

export { requireAuth };
