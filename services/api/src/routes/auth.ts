import type { FastifyInstance, FastifyReply, FastifyRequest } from "fastify";
import crypto from "node:crypto";
import { z } from "zod";
import { writeAuditEvent } from "../audit/audit";
import { generateTotpSecret, buildOtpAuthUrl, verifyTotpCode } from "../auth/mfa";
import { hashPassword, verifyPassword } from "../auth/password";
import { createSession, lookupSessionByToken, revokeSession } from "../auth/sessions";
import { withTransaction } from "../db/tx";
import { getClientIp, getUserAgent } from "../http/request-meta";

function extractSessionToken(request: FastifyRequest): string | null {
  const cookieName = request.server.config.sessionCookieName;
  const cookieToken = request.cookies?.[cookieName];
  if (cookieToken && typeof cookieToken === "string") return cookieToken;

  const auth = request.headers.authorization;
  if (!auth || typeof auth !== "string") return null;
  const [kind, token] = auth.split(" ");
  if (kind?.toLowerCase() !== "bearer") return null;
  return token ?? null;
}

async function requireAuth(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const token = extractSessionToken(request);
  if (!token) {
    reply.code(401).send({ error: "unauthorized" });
    return;
  }

  const found = await lookupSessionByToken(request.server.db, token);
  if (!found) {
    reply.code(401).send({ error: "unauthorized" });
    return;
  }

  request.user = found.user;
  request.session = found.session;
}

export function registerAuthRoutes(app: FastifyInstance): void {
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

    await writeAuditEvent(app.db, {
      orgId,
      userId,
      userEmail: email,
      eventType: "auth.login",
      resourceType: "session",
      resourceId: sessionId,
      sessionId,
      success: true,
      details: { method: "password", operation: "register" },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

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
      await writeAuditEvent(app.db, {
        userEmail: email,
        eventType: "auth.login_failed",
        resourceType: "user",
        resourceId: null,
        success: false,
        errorCode: "invalid_credentials",
        details: { method: "password" },
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      });
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
      await writeAuditEvent(app.db, {
        userId: row.id,
        userEmail: row.email,
        eventType: "auth.login_failed",
        resourceType: "user",
        resourceId: row.id,
        success: false,
        errorCode: "password_login_disabled",
        details: { method: "password" },
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      });
      return reply.code(401).send({ error: "invalid_credentials" });
    }

    const ok = await verifyPassword(password, row.password_hash);
    if (!ok) {
      await writeAuditEvent(app.db, {
        userId: row.id,
        userEmail: row.email,
        eventType: "auth.login_failed",
        resourceType: "user",
        resourceId: row.id,
        success: false,
        errorCode: "invalid_credentials",
        details: { method: "password" },
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      });
      return reply.code(401).send({ error: "invalid_credentials" });
    }

    if (row.mfa_totp_enabled) {
      const code = body.data.mfaCode;
      if (!code || !row.mfa_totp_secret || !verifyTotpCode(row.mfa_totp_secret, code)) {
        await writeAuditEvent(app.db, {
          userId: row.id,
          userEmail: row.email,
          eventType: "auth.login_failed",
          resourceType: "user",
          resourceId: row.id,
          success: false,
          errorCode: "mfa_required",
          details: { method: "password" },
          ipAddress: getClientIp(request),
          userAgent: getUserAgent(request)
        });
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

    await writeAuditEvent(app.db, {
      userId: row.id,
      userEmail: row.email,
      eventType: "auth.login",
      resourceType: "session",
      resourceId: sessionId,
      sessionId,
      success: true,
      details: { method: "password" },
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ user: { id: row.id, email: row.email, name: row.name } });
  });

  app.post("/auth/logout", { preHandler: requireAuth }, async (request, reply) => {
    const sessionId = request.session?.id;
    if (sessionId) {
      await revokeSession(app.db, sessionId);
      await writeAuditEvent(app.db, {
        userId: request.user?.id,
        userEmail: request.user?.email,
        eventType: "auth.logout",
        resourceType: "session",
        resourceId: sessionId,
        sessionId,
        success: true,
        ipAddress: getClientIp(request),
        userAgent: getUserAgent(request)
      });
    }

    reply.clearCookie(app.config.sessionCookieName, { path: "/" });
    return reply.send({ ok: true });
  });

  app.get("/me", { preHandler: requireAuth }, async (request) => {
    const orgs = await app.db.query(
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
      user: request.user,
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
    await writeAuditEvent(app.db, {
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "auth.mfa_enabled",
      resourceType: "user",
      resourceId: request.user!.id,
      sessionId: request.session?.id,
      success: true,
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

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

    await writeAuditEvent(app.db, {
      userId: request.user!.id,
      userEmail: request.user!.email,
      eventType: "auth.mfa_disabled",
      resourceType: "user",
      resourceId: request.user!.id,
      sessionId: request.session?.id,
      success: true,
      ipAddress: getClientIp(request),
      userAgent: getUserAgent(request)
    });

    return reply.send({ ok: true });
  });

  // OIDC / SSO hooks (scaffolding). Production code should implement provider-specific flows and callback verification.
  app.get("/auth/oidc/:provider/start", async (request, reply) => {
    const provider = (request.params as { provider: string }).provider;
    return reply.code(501).send({
      error: "not_implemented",
      message: `OIDC start not implemented for provider "${provider}" yet`
    });
  });

  app.get("/auth/oidc/:provider/callback", async (request, reply) => {
    const provider = (request.params as { provider: string }).provider;
    return reply.code(501).send({
      error: "not_implemented",
      message: `OIDC callback not implemented for provider "${provider}" yet`
    });
  });
}

export { requireAuth };
