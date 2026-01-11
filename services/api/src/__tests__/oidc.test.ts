import crypto from "node:crypto";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import jwt from "jsonwebtoken";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { deriveSecretStoreKey, putSecret } from "../secrets/secretStore";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // services/api/src/__tests__ -> services/api/migrations
  return path.resolve(here, "../../migrations");
}

function extractCookie(setCookieHeader: string | string[] | undefined): string {
  if (!setCookieHeader) throw new Error("missing set-cookie header");
  const raw = Array.isArray(setCookieHeader) ? setCookieHeader[0] : setCookieHeader;
  return raw.split(";")[0];
}

function parseJsonValue(value: unknown): any {
  if (!value) return null;
  if (typeof value === "object") return value;
  if (typeof value === "string") return JSON.parse(value);
  return null;
}

type CodeEntry = { nonce: string; sub: string; email: string; amr?: string[] };

class MockOidcProvider {
  private server: http.Server;
  private privateKey: crypto.KeyObject;
  private publicJwk: Record<string, unknown>;
  private readonly codes = new Map<string, CodeEntry>();

  readonly kid = "test-kid";
  readonly clientId = "test-client";
  readonly clientSecret = "test-client-secret";
  issuerUrl = "";

  constructor() {
    const { publicKey, privateKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
    this.privateKey = privateKey;
    const jwk = publicKey.export({ format: "jwk" }) as Record<string, unknown>;
    this.publicJwk = { ...jwk, kid: this.kid, use: "sig", alg: "RS256" };

    this.server = http.createServer((req, res) => {
      void this.handle(req, res);
    });
  }

  registerCode(code: string, entry: CodeEntry): void {
    this.codes.set(code, entry);
  }

  private async handle(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    const baseUrl = `http://${req.headers.host}`;
    const url = new URL(req.url ?? "/", baseUrl);

    if (url.pathname === "/.well-known/openid-configuration") {
      const issuer = baseUrl;
      const body = JSON.stringify({
        issuer,
        authorization_endpoint: `${issuer}/authorize`,
        token_endpoint: `${issuer}/token`,
        jwks_uri: `${issuer}/jwks`,
        response_types_supported: ["code"],
        subject_types_supported: ["public"],
        id_token_signing_alg_values_supported: ["RS256"]
      });
      res.writeHead(200, { "content-type": "application/json" });
      res.end(body);
      return;
    }

    if (url.pathname === "/jwks") {
      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify({ keys: [this.publicJwk] }));
      return;
    }

    if (url.pathname === "/token" && req.method === "POST") {
      const body = await new Promise<string>((resolve, reject) => {
        let data = "";
        req.on("data", (chunk) => (data += chunk.toString()));
        req.on("end", () => resolve(data));
        req.on("error", reject);
      });

      const params = new URLSearchParams(body);
      const code = params.get("code");
      const clientId = params.get("client_id");
      const clientSecret = params.get("client_secret");
      const codeVerifier = params.get("code_verifier");

      if (!code || !this.codes.has(code)) {
        res.writeHead(400, { "content-type": "application/json" });
        res.end(JSON.stringify({ error: "invalid_grant" }));
        return;
      }
      if (clientId !== this.clientId || clientSecret !== this.clientSecret) {
        res.writeHead(401, { "content-type": "application/json" });
        res.end(JSON.stringify({ error: "invalid_client" }));
        return;
      }
      if (!codeVerifier || codeVerifier.length < 10) {
        res.writeHead(400, { "content-type": "application/json" });
        res.end(JSON.stringify({ error: "invalid_request" }));
        return;
      }

      const entry = this.codes.get(code)!;
      const now = Math.floor(Date.now() / 1000);
      const issuer = baseUrl;
      const idToken = jwt.sign(
        {
          iss: issuer,
          aud: this.clientId,
          sub: entry.sub,
          email: entry.email,
          nonce: entry.nonce,
          amr: entry.amr,
          iat: now,
          exp: now + 60
        },
        this.privateKey,
        { algorithm: "RS256", keyid: this.kid }
      );

      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify({ access_token: "access", token_type: "Bearer", id_token: idToken }));
      return;
    }

    res.writeHead(404, { "content-type": "text/plain" });
    res.end("not found");
  }

  async start(): Promise<void> {
    await new Promise<void>((resolve) => {
      this.server.listen(0, "127.0.0.1", () => resolve());
    });
    const addr = this.server.address();
    if (!addr || typeof addr === "string") throw new Error("expected provider to listen on tcp port");
    this.issuerUrl = `http://127.0.0.1:${addr.port}`;
  }

  async stop(): Promise<void> {
    await new Promise<void>((resolve, reject) => {
      this.server.close((err) => (err ? reject(err) : resolve()));
    });
  }
}

async function createTestApp(): Promise<{
  db: Pool;
  config: AppConfig;
  app: ReturnType<typeof buildApp>;
}> {
  const mem = newDb({ autoCreateForeignKeyIndices: true });
  const pgAdapter = mem.adapters.createPg();
  const db = new pgAdapter.Pool();
  await runMigrations(db, { migrationsDir: getMigrationsDir() });

  const config: AppConfig = {
    port: 0,
    databaseUrl: "postgres://unused",
    sessionCookieName: "formula_session",
    sessionTtlSeconds: 60 * 60,
    cookieSecure: false,
    corsAllowedOrigins: [],
    syncTokenSecret: "test-sync-secret",
    syncTokenTtlSeconds: 60,
    secretStoreKeys: {
      currentKeyId: "legacy",
      keys: { legacy: deriveSecretStoreKey("test-secret-store-key") }
    },
    localKmsMasterKey: "test-local-kms-master-key",
    awsKmsEnabled: false,
    retentionSweepIntervalMs: null
  };

  const app = buildApp({ db, config });
  await app.ready();

  return { db, config, app };
}

describe("OIDC SSO", () => {
  const provider = new MockOidcProvider();

  beforeAll(async () => {
    await provider.start();
  });

  afterAll(async () => {
    await provider.stop();
  });

  it("successful login provisions user + membership and issues a session", async () => {
    const { db, config, app } = await createTestApp();
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "sso-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "SSO Org"
        }
      });
      expect(ownerRegister.statusCode).toBe(200);
      const orgId = (ownerRegister.json() as any).organization.id as string;

      await db.query("UPDATE org_settings SET allowed_auth_methods = $2::jsonb WHERE org_id = $1", [
        orgId,
        JSON.stringify(["password", "oidc"])
      ]);

      await db.query(
        `
          INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
          VALUES ($1,$2,$3,$4,$5::jsonb,$6)
        `,
        [orgId, "mock", provider.issuerUrl, provider.clientId, JSON.stringify(["openid", "email"]), true]
      );

      await putSecret(db, config.secretStoreKeys, `oidc:${orgId}:mock`, provider.clientSecret);

      const startRes = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/start`
      });
      expect(startRes.statusCode).toBe(302);
      const location = startRes.headers.location as string;
      const authUrl = new URL(location);
      expect(`${authUrl.origin}${authUrl.pathname}`).toBe(`${provider.issuerUrl}/authorize`);

      const state = authUrl.searchParams.get("state");
      const nonce = authUrl.searchParams.get("nonce");
      expect(state).toBeTruthy();
      expect(nonce).toBeTruthy();

      const code = "good-code";
      provider.registerCode(code, {
        nonce: nonce!,
        sub: "user-subject-123",
        email: "sso-user@example.com"
      });

      const callbackRes = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/callback?code=${encodeURIComponent(code)}&state=${encodeURIComponent(
          state!
        )}`
      });
      expect(callbackRes.statusCode).toBe(200);
      const cookie = extractCookie(callbackRes.headers["set-cookie"]);
      expect(cookie.startsWith(`${config.sessionCookieName}=`)).toBe(true);

      const me = await app.inject({
        method: "GET",
        url: "/me",
        headers: { cookie }
      });
      expect(me.statusCode).toBe(200);
      const meBody = me.json() as any;
      expect(meBody.user.email).toBe("sso-user@example.com");
      expect(meBody.organizations.some((o: any) => o.id === orgId)).toBe(true);

      const member = await db.query("SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2", [
        orgId,
        meBody.user.id
      ]);
      expect(member.rowCount).toBe(1);
      expect(member.rows[0].role).toBe("member");

      const identities = await db.query(
        "SELECT provider, subject, email, org_id FROM user_identities WHERE org_id = $1",
        [orgId]
      );
      expect(identities.rowCount).toBe(1);
      expect(identities.rows[0]).toMatchObject({
        provider: "mock",
        subject: "user-subject-123",
        email: "sso-user@example.com",
        org_id: orgId
      });

      const audit = await db.query(
        "SELECT event_type, org_id, details FROM audit_log WHERE event_type = 'auth.login' ORDER BY created_at DESC LIMIT 1"
      );
      expect(audit.rowCount).toBe(1);
      expect(audit.rows[0].org_id).toBe(orgId);
      expect(parseJsonValue(audit.rows[0].details)).toMatchObject({ method: "oidc", provider: "mock" });
    } finally {
      await app.close();
      await db.end();
    }
  });

  it("derives redirect_uri from forwarded headers only when TRUST_PROXY=true", async () => {
    const { db, config, app } = await createTestApp();
    let proxyApp: ReturnType<typeof buildApp> | null = null;
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "trust-proxy-oidc-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "Trust Proxy OIDC Org"
        }
      });
      expect(ownerRegister.statusCode).toBe(200);
      const orgId = (ownerRegister.json() as any).organization.id as string;

      await db.query("UPDATE org_settings SET allowed_auth_methods = $2::jsonb WHERE org_id = $1", [
        orgId,
        JSON.stringify(["password", "oidc"])
      ]);

      await db.query(
        `
          INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
          VALUES ($1,$2,$3,$4,$5::jsonb,$6)
        `,
        [orgId, "mock", provider.issuerUrl, provider.clientId, JSON.stringify(["openid", "email"]), true]
      );

      await putSecret(db, config.secretStoreKeys, `oidc:${orgId}:mock`, provider.clientSecret);

      const spoofedForwarded = {
        host: "good.example",
        "x-forwarded-host": "evil.example",
        "x-forwarded-proto": "https"
      };

      const startRes = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/start`,
        headers: spoofedForwarded
      });
      expect(startRes.statusCode).toBe(302);
      const startUrl = new URL(startRes.headers.location as string);
      expect(startUrl.searchParams.get("redirect_uri")).toBe(
        `http://good.example/auth/oidc/${orgId}/mock/callback`
      );

      proxyApp = buildApp({ db, config: { ...config, trustProxy: true } });
      await proxyApp.ready();

      const startResTrusted = await proxyApp.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/start`,
        headers: spoofedForwarded
      });
      expect(startResTrusted.statusCode).toBe(302);
      const trustedUrl = new URL(startResTrusted.headers.location as string);
      expect(trustedUrl.searchParams.get("redirect_uri")).toBe(
        `https://evil.example/auth/oidc/${orgId}/mock/callback`
      );
    } finally {
      await proxyApp?.close();
      await app.close();
      await db.end();
    }
  });

  it("uses PUBLIC_BASE_URL for OIDC redirect_uri (ignores Host / forwarded headers)", async () => {
    const { db, config, app } = await createTestApp();
    let publicApp: ReturnType<typeof buildApp> | null = null;
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "public-base-url-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "Public Base URL OIDC Org"
        }
      });
      expect(ownerRegister.statusCode).toBe(200);
      const orgId = (ownerRegister.json() as any).organization.id as string;

      await db.query("UPDATE org_settings SET allowed_auth_methods = $2::jsonb WHERE org_id = $1", [
        orgId,
        JSON.stringify(["password", "oidc"])
      ]);

      await db.query(
        `
          INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
          VALUES ($1,$2,$3,$4,$5::jsonb,$6)
        `,
        [orgId, "mock", provider.issuerUrl, provider.clientId, JSON.stringify(["openid", "email"]), true]
      );

      await putSecret(db, config.secretStoreKey, `oidc:${orgId}:mock`, provider.clientSecret);

      publicApp = buildApp({
        db,
        config: { ...config, publicBaseUrl: "https://api.public.example", trustProxy: true }
      });
      await publicApp.ready();

      const startRes = await publicApp.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/start`,
        headers: {
          host: "evil-host.example",
          "x-forwarded-host": "evil-forwarded.example",
          "x-forwarded-proto": "https"
        }
      });
      expect(startRes.statusCode).toBe(302);
      const authUrl = new URL(startRes.headers.location as string);
      expect(authUrl.searchParams.get("redirect_uri")).toBe(
        `https://api.public.example/auth/oidc/${orgId}/mock/callback`
      );
    } finally {
      await publicApp?.close();
      await app.close();
      await db.end();
    }
  });

  it("fails on invalid state", async () => {
    const { db, app } = await createTestApp();
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "invalid-state-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "Invalid State Org"
        }
      });
      const orgId = (ownerRegister.json() as any).organization.id as string;

      const callbackRes = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/callback?code=whatever&state=wrong`
      });
      expect(callbackRes.statusCode).toBe(401);
      expect((callbackRes.json() as any).error).toBe("invalid_state");
    } finally {
      await app.close();
      await db.end();
    }
  });

  it("fails on nonce mismatch", async () => {
    const { db, config, app } = await createTestApp();
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "nonce-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "Nonce Org"
        }
      });
      const orgId = (ownerRegister.json() as any).organization.id as string;

      await db.query("UPDATE org_settings SET allowed_auth_methods = $2::jsonb WHERE org_id = $1", [
        orgId,
        JSON.stringify(["password", "oidc"])
      ]);

      await db.query(
        `
          INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
          VALUES ($1,$2,$3,$4,$5::jsonb,$6)
        `,
        [orgId, "mock", provider.issuerUrl, provider.clientId, JSON.stringify(["openid", "email"]), true]
      );
      await putSecret(db, config.secretStoreKeys, `oidc:${orgId}:mock`, provider.clientSecret);

      const startRes = await app.inject({ method: "GET", url: `/auth/oidc/${orgId}/mock/start` });
      const authUrl = new URL(startRes.headers.location as string);
      const state = authUrl.searchParams.get("state")!;
      const nonce = authUrl.searchParams.get("nonce")!;

      const code = "bad-nonce-code";
      provider.registerCode(code, {
        nonce: `${nonce}-tampered`,
        sub: "user-subject-456",
        email: "nonce-user@example.com"
      });

      const callbackRes = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/callback?code=${encodeURIComponent(code)}&state=${encodeURIComponent(
          state
        )}`
      });
      expect(callbackRes.statusCode).toBe(401);
      expect((callbackRes.json() as any).error).toBe("invalid_nonce");
    } finally {
      await app.close();
      await db.end();
    }
  });

  it("blocks disabled providers", async () => {
    const { db, config, app } = await createTestApp();
    try {
      const ownerRegister = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "disabled-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "Disabled Provider Org"
        }
      });
      const orgId = (ownerRegister.json() as any).organization.id as string;

      await db.query("UPDATE org_settings SET allowed_auth_methods = $2::jsonb WHERE org_id = $1", [
        orgId,
        JSON.stringify(["password", "oidc"])
      ]);

      await db.query(
        `
          INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
          VALUES ($1,$2,$3,$4,$5::jsonb,$6)
        `,
        [orgId, "mock", provider.issuerUrl, provider.clientId, JSON.stringify(["openid", "email"]), false]
      );
      await putSecret(db, config.secretStoreKeys, `oidc:${orgId}:mock`, provider.clientSecret);

      const startRes = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/mock/start`
      });
      expect(startRes.statusCode).toBe(403);
      expect((startRes.json() as any).error).toBe("provider_disabled");
    } finally {
      await app.close();
      await db.end();
    }
  });
});
