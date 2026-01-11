import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";
import { deriveSecretStoreKey } from "../secrets/secretStore";

const SCIM_SCHEMA_LIST_RESPONSE = "urn:ietf:params:scim:api:messages:2.0:ListResponse";
const SCIM_SCHEMA_PATCH_OP = "urn:ietf:params:scim:api:messages:2.0:PatchOp";
const SCIM_SCHEMA_ERROR = "urn:ietf:params:scim:api:messages:2.0:Error";

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

async function registerOwner(app: ReturnType<typeof buildApp>, email: string, orgName: string) {
  const register = await app.inject({
    method: "POST",
    url: "/auth/register",
    payload: {
      email,
      password: "password1234",
      name: "Owner",
      orgName
    }
  });
  expect(register.statusCode).toBe(200);
  const cookie = extractCookie(register.headers["set-cookie"]);
  const orgId = (register.json() as any).organization.id as string;
  return { cookie, orgId };
}

async function createScimToken(app: ReturnType<typeof buildApp>, cookie: string, orgId: string, name: string) {
  const create = await app.inject({
    method: "POST",
    url: `/orgs/${orgId}/scim/tokens`,
    headers: { cookie },
    payload: { name }
  });
  expect(create.statusCode).toBe(200);
  const body = create.json() as any;
  return { tokenId: body.id as string, token: body.token as string };
}

describe("SCIM provisioning (Users)", () => {
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

  it("manages SCIM tokens (hashed storage; list hides token; revoked token rejected)", async () => {
    const { cookie, orgId } = await registerOwner(app, "scim-admin@example.com", "SCIM Org");

    const { tokenId, token } = await createScimToken(app, cookie, orgId, "okta");
    expect(token).toMatch(/^scim_[0-9a-f-]{36}\./i);

    const row = await db.query("SELECT token_hash FROM org_scim_tokens WHERE id = $1", [tokenId]);
    expect(row.rowCount).toBe(1);
    const tokenHash = row.rows[0]!.token_hash as string;
    const secret = token.split(".")[1]!;
    expect(tokenHash).not.toContain(secret);
    expect(tokenHash).not.toBe(token);

    const listTokens = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/scim/tokens`,
      headers: { cookie }
    });
    expect(listTokens.statusCode).toBe(200);
    const listBody = listTokens.json() as any;
    expect(listBody.tokens).toHaveLength(1);
    expect(listBody.tokens[0]).toMatchObject({ id: tokenId, orgId, name: "okta" });
    expect(listBody.tokens[0].token).toBeUndefined();

    const unauthorized = await app.inject({ method: "GET", url: "/scim/v2/Users" });
    expect(unauthorized.statusCode).toBe(401);
    expect((unauthorized.json() as any).schemas).toEqual([SCIM_SCHEMA_ERROR]);
    expect((unauthorized.json() as any).status).toBe("401");

    const listUsers = await app.inject({
      method: "GET",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${token}` }
    });
    expect(listUsers.statusCode).toBe(200);
    const usersBody = listUsers.json() as any;
    expect(usersBody.schemas).toEqual([SCIM_SCHEMA_LIST_RESPONSE]);
    expect(Array.isArray(usersBody.Resources)).toBe(true);
    expect(listUsers.headers["content-type"]).toContain("application/scim+json");

    const used = await db.query("SELECT last_used_at FROM org_scim_tokens WHERE id = $1", [tokenId]);
    expect(used.rows[0]!.last_used_at).toBeTruthy();

    const revoke = await app.inject({
      method: "DELETE",
      url: `/orgs/${orgId}/scim/tokens/${tokenId}`,
      headers: { cookie }
    });
    expect(revoke.statusCode).toBe(200);

    const revokedList = await app.inject({
      method: "GET",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${token}` }
    });
    expect(revokedList.statusCode).toBe(401);
    expect((revokedList.json() as any).schemas).toEqual([SCIM_SCHEMA_ERROR]);

    const audit = await db.query(
      "SELECT event_type FROM audit_log WHERE resource_id = $1 ORDER BY created_at ASC",
      [tokenId]
    );
    const eventTypes = audit.rows.map((r) => r.event_type as string);
    expect(eventTypes).toContain("org.scim_token.created");
    expect(eventTypes).toContain("org.scim_token.revoked");
  });

  it(
    "creates users, toggles active, and scopes list responses to the token org",
    async () => {
    const { cookie: cookieA, orgId: orgA } = await registerOwner(app, "org-a-owner@example.com", "Org A");
    const { cookie: cookieB, orgId: orgB } = await registerOwner(app, "org-b-owner@example.com", "Org B");

    const { token: tokenA } = await createScimToken(app, cookieA, orgA, "a");
    const { token: tokenB } = await createScimToken(app, cookieB, orgB, "b");

    const createA = await app.inject({
      method: "POST",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${tokenA}` },
      payload: {
        schemas: ["urn:ietf:params:scim:schemas:core:2.0:User"],
        userName: "alice@example.com",
        name: { givenName: "Alice", familyName: "Example" },
        active: true
      }
    });
    expect(createA.statusCode).toBe(201);
    const userAId = (createA.json() as any).id as string;

    const createB = await app.inject({
      method: "POST",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${tokenB}` },
      payload: {
        schemas: ["urn:ietf:params:scim:schemas:core:2.0:User"],
        userName: "bob@example.com",
        name: { givenName: "Bob", familyName: "Example" },
        active: true
      }
    });
    expect(createB.statusCode).toBe(201);
    const userBId = (createB.json() as any).id as string;

    const memberA = await db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [orgA, userAId]);
    expect(memberA.rowCount).toBe(1);

    const deactivate = await app.inject({
      method: "PATCH",
      url: `/scim/v2/Users/${userAId}`,
      headers: { authorization: `Bearer ${tokenA}` },
      payload: {
        schemas: [SCIM_SCHEMA_PATCH_OP],
        Operations: [{ op: "Replace", path: "active", value: false }]
      }
    });
    expect(deactivate.statusCode).toBe(200);

    const memberAfterDeactivate = await db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [
      orgA,
      userAId
    ]);
    expect(memberAfterDeactivate.rowCount).toBe(0);

    const reactivate = await app.inject({
      method: "PATCH",
      url: `/scim/v2/Users/${userAId}`,
      headers: { authorization: `Bearer ${tokenA}` },
      payload: {
        schemas: [SCIM_SCHEMA_PATCH_OP],
        Operations: [{ op: "Replace", path: "active", value: true }]
      }
    });
    expect(reactivate.statusCode).toBe(200);

    const memberAfterReactivate = await db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [
      orgA,
      userAId
    ]);
    expect(memberAfterReactivate.rowCount).toBe(1);

    const listA = await app.inject({
      method: "GET",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${tokenA}` }
    });
    expect(listA.statusCode).toBe(200);
    const bodyA = listA.json() as any;
    const userNamesA = (bodyA.Resources as any[]).map((r) => r.userName);
    expect(userNamesA).toContain("alice@example.com");
    expect(userNamesA).not.toContain("bob@example.com");

    const listB = await app.inject({
      method: "GET",
      url: "/scim/v2/Users",
      headers: { authorization: `Bearer ${tokenB}` }
    });
    expect(listB.statusCode).toBe(200);
    const bodyB = listB.json() as any;
    const userNamesB = (bodyB.Resources as any[]).map((r) => r.userName);
    expect(userNamesB).toContain("bob@example.com");
    expect(userNamesB).not.toContain("alice@example.com");

    const auditA = await db.query(
      "SELECT event_type, details FROM audit_log WHERE resource_id = $1 ORDER BY created_at ASC",
      [userAId]
    );
    const eventTypesA = auditA.rows.map((r) => r.event_type as string);
    expect(eventTypesA).toContain("scim.user.created");
    expect(eventTypesA).toContain("scim.user.deactivated");
    expect(eventTypesA).toContain("scim.user.reactivated");
    for (const row of auditA.rows) {
      expect((row.details as any)?.source).toBe("scim");
    }

    // Ensure org B cannot delete org A membership by id.
    const deleteCross = await app.inject({
      method: "DELETE",
      url: `/scim/v2/Users/${userAId}`,
      headers: { authorization: `Bearer ${tokenB}` }
    });
    expect(deleteCross.statusCode).toBe(404);

    // Cleanup assertion: org B membership exists.
    const memberB = await db.query("SELECT 1 FROM org_members WHERE org_id = $1 AND user_id = $2", [orgB, userBId]);
    expect(memberB.rowCount).toBe(1);
    },
    15_000
  );
});
