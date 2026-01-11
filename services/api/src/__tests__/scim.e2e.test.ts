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

describe("SCIM provisioning", () => {
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
      oidcAuthStateCleanupIntervalMs: null
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("provisions users via SCIM token and logs audit events", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "scim-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "SCIM Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const createToken = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/scim/token`,
      headers: { cookie }
    });
    expect(createToken.statusCode).toBe(200);
    const token = (createToken.json() as any).token as string;
    expect(token).toMatch(new RegExp(`^scim_${orgId}\\.`));

    const createAlice = await app.inject({
      method: "POST",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${token}` },
      payload: {
        schemas: ["urn:ietf:params:scim:schemas:core:2.0:User"],
        userName: "alice@example.com",
        displayName: "Alice",
        emails: [{ value: "alice@example.com", primary: true }],
        active: true
      }
    });
    expect(createAlice.statusCode).toBe(201);
    const alice = createAlice.json() as any;
    const aliceId = alice.id as string;

    const userRow = await db.query("SELECT id, email, name FROM users WHERE email = $1", ["alice@example.com"]);
    expect(userRow.rowCount).toBe(1);
    expect(userRow.rows[0]).toMatchObject({ id: aliceId, email: "alice@example.com", name: "Alice" });

    const memberRow = await db.query("SELECT role FROM org_members WHERE org_id = $1 AND user_id = $2", [
      orgId,
      aliceId
    ]);
    expect(memberRow.rowCount).toBe(1);
    expect(memberRow.rows[0]!.role).toBe("member");

    await app.inject({
      method: "POST",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${token}` },
      payload: {
        schemas: ["urn:ietf:params:scim:schemas:core:2.0:User"],
        userName: "bob@example.com",
        displayName: "Bob",
        emails: [{ value: "bob@example.com", primary: true }],
        active: true
      }
    });

    const listPage1 = await app.inject({
      method: "GET",
      url: "/scim/v2/Users?startIndex=1&count=1",
      headers: { authorization: `Bearer ${token}` }
    });
    expect(listPage1.statusCode).toBe(200);
    const page1 = listPage1.json() as any;
    // Includes the org owner created via /auth/register.
    expect(page1.totalResults).toBe(3);
    expect(page1.itemsPerPage).toBe(1);
    expect(page1.Resources).toHaveLength(1);

    const listPage2 = await app.inject({
      method: "GET",
      url: "/scim/v2/Users?startIndex=2&count=1",
      headers: { authorization: `Bearer ${token}` }
    });
    expect(listPage2.statusCode).toBe(200);
    const page2 = listPage2.json() as any;
    expect(page2.totalResults).toBe(3);
    expect(page2.itemsPerPage).toBe(1);
    expect(page2.Resources).toHaveLength(1);

    const filterAlice = await app.inject({
      method: "GET",
      url: `/scim/v2/Users?filter=${encodeURIComponent('userName eq "alice@example.com"')}`,
      headers: { authorization: `Bearer ${token}` }
    });
    expect(filterAlice.statusCode).toBe(200);
    const filterBody = filterAlice.json() as any;
    expect(filterBody.totalResults).toBe(1);
    expect(filterBody.Resources).toHaveLength(1);
    expect(filterBody.Resources[0].userName).toBe("alice@example.com");

    const deactivate = await app.inject({
      method: "PATCH",
      url: `/scim/v2/Users/${aliceId}`,
      headers: { authorization: `Bearer ${token}` },
      payload: {
        schemas: ["urn:ietf:params:scim:api:messages:2.0:PatchOp"],
        Operations: [{ op: "Replace", path: "active", value: false }]
      }
    });
    expect(deactivate.statusCode).toBe(200);
    expect((deactivate.json() as any).active).toBe(false);

    const removed = await db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [orgId, aliceId]);
    expect(removed.rowCount).toBe(0);

    const reactivate = await app.inject({
      method: "PATCH",
      url: `/scim/v2/Users/${aliceId}`,
      headers: { authorization: `Bearer ${token}` },
      payload: {
        schemas: ["urn:ietf:params:scim:api:messages:2.0:PatchOp"],
        Operations: [{ op: "Replace", path: "active", value: true }]
      }
    });
    expect(reactivate.statusCode).toBe(200);
    expect((reactivate.json() as any).active).toBe(true);

    const readded = await db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [orgId, aliceId]);
    expect(readded.rowCount).toBe(1);

    const auditRes = await db.query(
      "SELECT event_type, details FROM audit_log WHERE org_id = $1 ORDER BY created_at ASC",
      [orgId]
    );
    const eventTypes = auditRes.rows.map((row) => row.event_type as string);
    expect(eventTypes).toContain("org.scim.token_created");
    expect(eventTypes).toContain("admin.user_created");
    expect(eventTypes).toContain("admin.user_deactivated");
    expect(eventTypes).toContain("admin.user_reactivated");

    const scimEvents = auditRes.rows.filter((row) => String(row.event_type).startsWith("admin.user_"));
    for (const row of scimEvents) {
      const details = row.details as any;
      const parsed = typeof details === "string" ? JSON.parse(details) : details;
      expect(parsed.source).toBe("scim");
    }

    const revokeToken = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/scim/token`,
      headers: { cookie }
    });
    expect(revokeToken.statusCode).toBe(200);

    const blocked = await app.inject({
      method: "GET",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${token}` }
    });
    expect(blocked.statusCode).toBe(401);

    const rotateToken = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/scim/token`,
      headers: { cookie }
    });
    expect(rotateToken.statusCode).toBe(200);
    const newToken = (rotateToken.json() as any).token as string;
    expect(newToken).not.toBe(token);

    const oldTokenStillBlocked = await app.inject({
      method: "GET",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${token}` }
    });
    expect(oldTokenStillBlocked.statusCode).toBe(401);

    const allowed = await app.inject({
      method: "GET",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${newToken}` }
    });
    expect(allowed.statusCode).toBe(200);
  });
});
