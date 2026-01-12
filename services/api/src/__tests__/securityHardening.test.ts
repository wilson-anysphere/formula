import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "../app";
import { loadConfig, type AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { deriveSecretStoreKey } from "../secrets/secretStore";

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

describe("security hardening", () => {
  let db: Pool;
  let config: AppConfig;
  let app: ReturnType<typeof buildApp>;
  let secureApp: ReturnType<typeof buildApp>;
  let trustProxyApp: ReturnType<typeof buildApp>;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

    config = {
      port: 0,
      databaseUrl: "postgres://unused",
      publicBaseUrl: "http://localhost",
      publicBaseUrlHostAllowlist: ["localhost"],
      trustProxy: false,
      sessionCookieName: "formula_session",
      sessionTtlSeconds: 60 * 60,
      cookieSecure: false,
      corsAllowedOrigins: ["https://allowed.example"],
      syncTokenSecret: "test-sync-secret",
      syncTokenTtlSeconds: 60,
      secretStoreKeys: {
        currentKeyId: "legacy",
        keys: { legacy: deriveSecretStoreKey("test-secret-store-key") }
      },
      localKmsMasterKey: "test-local-kms-master-key",
      awsKmsEnabled: false,
      retentionSweepIntervalMs: null,
      oidcAuthStateCleanupIntervalMs: null
    };

    app = buildApp({ db, config });
    secureApp = buildApp({ db, config: { ...config, cookieSecure: true } });
    trustProxyApp = buildApp({ db, config: { ...config, trustProxy: true } });
    await app.ready();
    await secureApp.ready();
    await trustProxyApp.ready();
  });

  afterAll(async () => {
    await trustProxyApp.close();
    await secureApp.close();
    await app.close();
    await db.end();
  });

  it("rate limits repeated login attempts and returns Retry-After", async () => {
    let limited: any = null;
    const limit = 20;
    for (let i = 0; i < limit + 10; i++) {
      const res = await app.inject({
        method: "POST",
        url: "/auth/login",
        remoteAddress: "198.51.100.10",
        payload: { email: "does-not-exist@example.com", password: "wrong-password" }
      });
      expect(res.headers["x-ratelimit-limit"]).toBe(String(limit));

      if (res.statusCode === 429) {
        limited = res;
        break;
      }
    }

    expect(limited).toBeTruthy();
    expect(limited.statusCode).toBe(429);
    expect(limited.headers["retry-after"]).toBeTypeOf("string");
    expect(Number(limited.headers["retry-after"])).toBeGreaterThan(0);
    expect((limited.json() as any).error).toBe("too_many_requests");
  });

  it("rate limits repeated registration attempts and returns Retry-After", async () => {
    let limited: any = null;
    const limit = 20;
    const payload = {
      email: "rate-limited-register@example.com",
      password: "password1234",
      name: "Register Rate Limit"
    };

    for (let i = 0; i < limit + 20; i++) {
      const res = await app.inject({
        method: "POST",
        url: "/auth/register",
        remoteAddress: "198.51.100.30",
        payload
      });
      expect(res.headers["x-ratelimit-limit"]).toBe(String(limit));

      if (res.statusCode === 429) {
        limited = res;
        break;
      }
    }

    expect(limited).toBeTruthy();
    expect(limited.statusCode).toBe(429);
    expect(limited.headers["retry-after"]).toBeTypeOf("string");
    expect(Number(limited.headers["retry-after"])).toBeGreaterThan(0);
    expect((limited.json() as any).error).toBe("too_many_requests");
  });

  it("rate limits repeated OIDC start requests and returns Retry-After", async () => {
    let limited: any = null;
    const orgId = "00000000-0000-0000-0000-000000000000";
    for (let i = 0; i < 50; i++) {
      const res = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/test/start`,
        remoteAddress: "198.51.100.20"
      });

      if (res.statusCode === 429) {
        limited = res;
        break;
      }
    }

    expect(limited).toBeTruthy();
    expect(limited.statusCode).toBe(429);
    expect(limited.headers["retry-after"]).toBeTypeOf("string");
    expect(Number(limited.headers["retry-after"])).toBeGreaterThan(0);
    expect((limited.json() as any).error).toBe("too_many_requests");
  });

  it("rate limits repeated OIDC callback requests and returns Retry-After", async () => {
    let limited: any = null;
    const orgId = "00000000-0000-0000-0000-000000000000";
    for (let i = 0; i < 50; i++) {
      const res = await app.inject({
        method: "GET",
        url: `/auth/oidc/${orgId}/test/callback`,
        remoteAddress: "198.51.100.21"
      });

      if (res.statusCode === 429) {
        limited = res;
        break;
      }
    }

    expect(limited).toBeTruthy();
    expect(limited.statusCode).toBe(429);
    expect(limited.headers["retry-after"]).toBeTypeOf("string");
    expect(Number(limited.headers["retry-after"])).toBeGreaterThan(0);
    expect((limited.json() as any).error).toBe("too_many_requests");
  });

  it("rate limits repeated SAML start requests and returns Retry-After", async () => {
    let limited: any = null;
    const orgId = "00000000-0000-0000-0000-000000000000";
    for (let i = 0; i < 50; i++) {
      const res = await app.inject({
        method: "GET",
        url: `/auth/saml/${orgId}/test/start`,
        remoteAddress: "198.51.100.11"
      });

      if (res.statusCode === 429) {
        limited = res;
        break;
      }
    }

    expect(limited).toBeTruthy();
    expect(limited.statusCode).toBe(429);
    expect(limited.headers["retry-after"]).toBeTypeOf("string");
    expect(Number(limited.headers["retry-after"])).toBeGreaterThan(0);
    expect((limited.json() as any).error).toBe("too_many_requests");
  });

  it("rate limits repeated SAML callback requests and returns Retry-After", async () => {
    let limited: any = null;
    const orgId = "00000000-0000-0000-0000-000000000000";
    for (let i = 0; i < 50; i++) {
      const res = await app.inject({
        method: "POST",
        url: `/auth/saml/${orgId}/test/callback`,
        headers: { "content-type": "application/x-www-form-urlencoded" },
        remoteAddress: "198.51.100.12",
        payload: "SAMLResponse=Zm9v"
      });

      if (res.statusCode === 429) {
        limited = res;
        break;
      }
    }

    expect(limited).toBeTruthy();
    expect(limited.statusCode).toBe(429);
    expect(limited.headers["retry-after"]).toBeTypeOf("string");
    expect(Number(limited.headers["retry-after"])).toBeGreaterThan(0);
    expect((limited.json() as any).error).toBe("too_many_requests");
  });

  it("sets baseline security headers (and HSTS when cookieSecure=true)", async () => {
    const res = await app.inject({ method: "GET", url: "/health" });
    expect(res.statusCode).toBe(200);
    expect(res.headers["server"]).toBeUndefined();
    expect(res.headers["x-dns-prefetch-control"]).toBe("off");
    expect(res.headers["x-download-options"]).toBe("noopen");
    expect(res.headers["x-content-type-options"]).toBe("nosniff");
    expect(res.headers["x-frame-options"]).toBe("DENY");
    expect(res.headers["referrer-policy"]).toBe("no-referrer");
    expect(res.headers["x-permitted-cross-domain-policies"]).toBe("none");
    expect(res.headers["x-robots-tag"]).toBe("noindex");
    expect(res.headers["content-security-policy"]).toContain("default-src 'none'");
    expect(res.headers["permissions-policy"]).toContain("camera=()");
    expect(res.headers["cache-control"]).toBe("no-store");
    expect(res.headers["strict-transport-security"]).toBeUndefined();

    const resSecure = await secureApp.inject({ method: "GET", url: "/health" });
    expect(resSecure.statusCode).toBe(200);
    expect(resSecure.headers["server"]).toBeUndefined();
    expect(resSecure.headers["x-dns-prefetch-control"]).toBe("off");
    expect(resSecure.headers["x-download-options"]).toBe("noopen");
    expect(resSecure.headers["x-permitted-cross-domain-policies"]).toBe("none");
    expect(resSecure.headers["x-robots-tag"]).toBe("noindex");
    expect(resSecure.headers["content-security-policy"]).toContain("default-src 'none'");
    expect(resSecure.headers["cache-control"]).toBe("no-store");
    expect(resSecure.headers["strict-transport-security"]).toContain("max-age=");
  });

  it("enforces CORS allowlist (credentials only for trusted origins)", async () => {
    const allowed = await app.inject({
      method: "GET",
      url: "/health",
      headers: {
        origin: "https://allowed.example"
      }
    });
    expect(allowed.statusCode).toBe(200);
    expect(allowed.headers["access-control-allow-origin"]).toBe("https://allowed.example");
    expect(allowed.headers["access-control-allow-credentials"]).toBe("true");

    const disallowed = await app.inject({
      method: "GET",
      url: "/health",
      headers: {
        origin: "https://evil.example"
      }
    });
    expect(disallowed.statusCode).toBe(200);
    expect(disallowed.headers["access-control-allow-origin"]).toBeUndefined();
    expect(disallowed.headers["access-control-allow-credentials"]).toBeUndefined();
  });

  it("validates production config and rejects dev defaults / COOKIE_SECURE!=true", () => {
    expect(() =>
      loadConfig({
        NODE_ENV: "production",
        PUBLIC_BASE_URL: "https://api.example.com",
        COOKIE_SECURE: "true"
      })
    ).toThrow(/default development secrets/i);

    expect(() =>
      loadConfig({
        NODE_ENV: "production",
        PUBLIC_BASE_URL: "https://api.example.com",
        COOKIE_SECURE: "true",
        SYNC_TOKEN_SECRET: "prod-sync-token-secret",
        SECRET_STORE_KEY: "prod-secret-store-key"
      })
    ).not.toThrow();

    expect(() =>
      loadConfig({
        NODE_ENV: "production",
        PUBLIC_BASE_URL: "https://api.example.com",
        COOKIE_SECURE: "true",
        SYNC_TOKEN_SECRET: "prod-sync-token-secret",
        SECRET_STORE_KEYS: `k1:${Buffer.alloc(32, 0x11).toString("base64")}`
      })
    ).not.toThrow();

    expect(() =>
      loadConfig({
        NODE_ENV: "production",
        PUBLIC_BASE_URL: "https://api.example.com",
        COOKIE_SECURE: "true",
        SYNC_TOKEN_SECRET: "prod-sync-token-secret",
        SECRET_STORE_KEYS_JSON: JSON.stringify({
          currentKeyId: "dev",
          keys: { dev: deriveSecretStoreKey("dev-secret-store-key-change-me").toString("base64") }
        })
      })
    ).toThrow(/default development secrets/i);

    expect(() =>
      loadConfig({
        NODE_ENV: "production",
        PUBLIC_BASE_URL: "https://api.example.com",
        COOKIE_SECURE: "true",
        SYNC_TOKEN_SECRET: "prod-sync-token-secret",
        SECRET_STORE_KEYS: `dev:${deriveSecretStoreKey("dev-secret-store-key-change-me").toString("base64")}`
      })
    ).toThrow(/default development secrets/i);

    expect(() =>
      loadConfig({
        NODE_ENV: "production",
        PUBLIC_BASE_URL: "https://api.example.com",
        COOKIE_SECURE: "false",
        SYNC_TOKEN_SECRET: "prod-sync-token-secret",
        SECRET_STORE_KEY: "prod-secret-store-key",
        LOCAL_KMS_MASTER_KEY: "prod-local-kms-master-key"
      })
    ).toThrow(/COOKIE_SECURE/i);
  });

  it("validates SECRET_STORE_KEYS_JSON schema", () => {
    const keyB64 = deriveSecretStoreKey("test-key").toString("base64");

    expect(() => loadConfig({ SECRET_STORE_KEYS_JSON: "not-json" })).toThrow(/valid JSON/i);

    expect(() =>
      loadConfig({
        SECRET_STORE_KEYS_JSON: JSON.stringify({ currentKeyId: "k1", keys: {} })
      })
    ).toThrow(/missing from keys/i);

    expect(() =>
      loadConfig({
        SECRET_STORE_KEYS_JSON: JSON.stringify({
          currentKeyId: "bad:key",
          keys: { "bad:key": keyB64 }
        })
      })
    ).toThrow(/must not contain/i);

    expect(() =>
      loadConfig({
        SECRET_STORE_KEYS_JSON: JSON.stringify({
          currentKeyId: "k1",
          keys: { "": keyB64, k1: keyB64 }
        })
      })
    ).toThrow(/must not contain empty key ids/i);
  });

  it("enforces org ip_allowlist for session auth on /orgs/:orgId/* and /docs/:docId/*", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "ip-allowlist-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Allowlist Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Allowlist Doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const setAllowlist = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { ipAllowlist: ["10.0.0.0/8"] }
    });
    expect(setAllowlist.statusCode).toBe(200);

    const blockedCreateDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      remoteAddress: "203.0.113.10",
      payload: { orgId, title: "Blocked Doc" }
    });
    expect(blockedCreateDoc.statusCode).toBe(403);
    expect((blockedCreateDoc.json() as any).error).toBe("ip_not_allowed");

    const blockedOrg = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}`,
      headers: { cookie },
      remoteAddress: "203.0.113.10"
    });
    expect(blockedOrg.statusCode).toBe(403);
    expect((blockedOrg.json() as any).error).toBe("ip_not_allowed");

    const blockedSiem = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      remoteAddress: "203.0.113.10"
    });
    expect(blockedSiem.statusCode).toBe(403);
    expect((blockedSiem.json() as any).error).toBe("ip_not_allowed");

    const blockedOidcProviders = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/oidc-providers`,
      headers: { cookie },
      remoteAddress: "203.0.113.10"
    });
    expect(blockedOidcProviders.statusCode).toBe(403);
    expect((blockedOidcProviders.json() as any).error).toBe("ip_not_allowed");

    const blockedSamlProviders = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/saml-providers`,
      headers: { cookie },
      remoteAddress: "203.0.113.10"
    });
    expect(blockedSamlProviders.statusCode).toBe(403);
    expect((blockedSamlProviders.json() as any).error).toBe("ip_not_allowed");

    const blockedDoc = await app.inject({
      method: "GET",
      url: `/docs/${docId}`,
      headers: { cookie },
      remoteAddress: "203.0.113.10"
    });
    expect(blockedDoc.statusCode).toBe(403);
    expect((blockedDoc.json() as any).error).toBe("ip_not_allowed");

    const blockedDocDlp = await app.inject({
      method: "POST",
      url: `/docs/${docId}/dlp/evaluate`,
      headers: { cookie },
      remoteAddress: "203.0.113.10",
      payload: { action: "export.csv" }
    });
    expect(blockedDocDlp.statusCode).toBe(403);
    expect((blockedDocDlp.json() as any).error).toBe("ip_not_allowed");

    const createShareLink = await app.inject({
      method: "POST",
      url: `/docs/${docId}/share-links`,
      headers: { cookie },
      remoteAddress: "10.1.2.3",
      payload: { visibility: "public", role: "viewer" }
    });
    expect(createShareLink.statusCode).toBe(200);
    const shareToken = (createShareLink.json() as any).shareLink.token as string;
    expect(shareToken).toBeTypeOf("string");

    const blockedRedeem = await app.inject({
      method: "POST",
      url: `/share-links/${shareToken}/redeem`,
      headers: { cookie },
      remoteAddress: "203.0.113.10"
    });
    expect(blockedRedeem.statusCode).toBe(403);
    expect((blockedRedeem.json() as any).error).toBe("ip_not_allowed");

    const allowedOrg = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}`,
      headers: { cookie },
      remoteAddress: "10.1.2.3"
    });
    expect(allowedOrg.statusCode).toBe(200);

    const audit = await db.query("SELECT event_type FROM audit_log WHERE event_type = $1", [
      "org.ip_allowlist.blocked"
    ]);
    expect(audit.rowCount).toBeGreaterThan(0);
  });

  it("uses X-Forwarded-For for allowlist enforcement when TRUST_PROXY=true", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "trust-proxy-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Trust Proxy Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const setAllowlist = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: { ipAllowlist: ["203.0.113.0/24"] }
    });
    expect(setAllowlist.statusCode).toBe(200);

    const blockedWithoutTrustProxy = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}`,
      headers: { cookie, "x-forwarded-for": "203.0.113.10" },
      remoteAddress: "10.0.0.1"
    });
    expect(blockedWithoutTrustProxy.statusCode).toBe(403);
    expect((blockedWithoutTrustProxy.json() as any).error).toBe("ip_not_allowed");

    const allowedWithTrustProxy = await trustProxyApp.inject({
      method: "GET",
      url: `/orgs/${orgId}`,
      headers: { cookie, "x-forwarded-for": "203.0.113.10" },
      remoteAddress: "10.0.0.1"
    });
    expect(allowedWithTrustProxy.statusCode).toBe(200);
  });

  it("rejects invalid orgId/docId path params on org-scoped routes", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "invalid-orgid-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Invalid OrgId Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const res = await app.inject({
      method: "GET",
      url: "/orgs/not-a-uuid",
      headers: { cookie }
    });
    expect(res.statusCode).toBe(400);
    expect((res.json() as any).error).toBe("invalid_request");

    const docRes = await app.inject({
      method: "GET",
      url: "/docs/not-a-uuid",
      headers: { cookie }
    });
    expect(docRes.statusCode).toBe(404);
    expect((docRes.json() as any).error).toBe("doc_not_found");

    const dlpRes = await app.inject({
      method: "POST",
      url: "/docs/not-a-uuid/dlp/evaluate",
      headers: { cookie },
      payload: { action: "export.csv" }
    });
    expect(dlpRes.statusCode).toBe(404);
    expect((dlpRes.json() as any).error).toBe("doc_not_found");

    const createDoc = await app.inject({
      method: "POST",
      url: "/docs",
      headers: { cookie },
      payload: { orgId, title: "Param validation doc" }
    });
    expect(createDoc.statusCode).toBe(200);
    const docId = (createDoc.json() as any).document.id as string;

    const versionRes = await app.inject({
      method: "GET",
      url: `/docs/${docId}/versions/not-a-uuid`,
      headers: { cookie }
    });
    expect(versionRes.statusCode).toBe(404);
    expect((versionRes.json() as any).error).toBe("version_not_found");

    const permissionRes = await app.inject({
      method: "DELETE",
      url: `/docs/${docId}/range-permissions/not-a-uuid`,
      headers: { cookie }
    });
    expect(permissionRes.statusCode).toBe(404);
    expect((permissionRes.json() as any).error).toBe("not_found");
  });
});
