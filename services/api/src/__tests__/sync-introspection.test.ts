import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { buildApp } from "../app";
import type { AppConfig } from "../config";
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

describe("internal: sync token introspection", () => {
  let db: Pool;
  let config: AppConfig;
  let app: ReturnType<typeof buildApp>;

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
      corsAllowedOrigins: [],
      syncTokenSecret: "test-sync-secret",
      syncTokenTtlSeconds: 60,
      secretStoreKeys: {
        currentKeyId: "legacy",
        keys: { legacy: deriveSecretStoreKey("test-secret-store-key") }
      },
      localKmsMasterKey: "test-local-kms-master-key",
      awsKmsEnabled: false,
      retentionSweepIntervalMs: null,
      oidcAuthStateCleanupIntervalMs: null,
      internalAdminToken: "internal-admin-token"
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it(
    "enforces org IP allowlist when clientIp is provided",
    async () => {
      const suffix = Math.random().toString(16).slice(2);
      const email = `introspect-ip-${suffix}@example.com`;

      const register = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email,
          password: "password1234",
          name: "User",
          orgName: "Org"
        }
      });
      expect(register.statusCode).toBe(200);
      const cookie = extractCookie(register.headers["set-cookie"]);
      const body = register.json() as any;
      const userId = body.user.id as string;
      const orgId = body.organization.id as string;

      const createDoc = await app.inject({
        method: "POST",
        url: "/docs",
        headers: { cookie },
        payload: { orgId, title: "Doc" }
      });
      expect(createDoc.statusCode).toBe(200);
      const docId = (createDoc.json() as any).document.id as string;

      const tokenRes = await app.inject({
        method: "POST",
        url: `/docs/${docId}/sync-token`,
        headers: { cookie }
      });
      expect(tokenRes.statusCode).toBe(200);
      const token = (tokenRes.json() as any).token as string;

      const setAllowlist = await app.inject({
        method: "PATCH",
        url: `/orgs/${orgId}/settings`,
        headers: { cookie },
        payload: { ipAllowlist: ["10.0.0.0/8"] }
      });
      expect(setAllowlist.statusCode).toBe(200);

      const allowed = await app.inject({
        method: "POST",
        url: "/internal/sync/introspect",
        headers: { "x-internal-admin-token": config.internalAdminToken! },
        payload: { token, docId, clientIp: "10.1.2.3", userAgent: "vitest" }
      });
      expect(allowed.statusCode).toBe(200);
      expect(allowed.json()).toMatchObject({ ok: true, userId, orgId, role: "owner" });

      const blocked = await app.inject({
        method: "POST",
        url: "/internal/sync/introspect",
        headers: { "x-internal-admin-token": config.internalAdminToken! },
        payload: { token, docId, clientIp: "203.0.113.5", userAgent: "vitest" }
      });
      expect(blocked.statusCode).toBe(200);
      expect(blocked.json()).toMatchObject({
        ok: false,
        active: false,
        error: "forbidden",
        reason: "ip_not_allowed"
      });
    },
    20_000
  );
});
