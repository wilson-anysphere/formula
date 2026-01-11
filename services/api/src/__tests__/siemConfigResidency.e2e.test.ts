import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "../app";
import type { AppConfig } from "../config";
import { runMigrations } from "../db/migrations";

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

describe("SIEM config data residency", () => {
  let db: Pool;
  let app: ReturnType<typeof buildApp>;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

    const config: AppConfig = {
      port: 0,
      databaseUrl: "postgres://unused",
      sessionCookieName: "formula_session",
      sessionTtlSeconds: 60 * 60,
      cookieSecure: false,
      syncTokenSecret: "test-sync-secret",
      syncTokenTtlSeconds: 60,
      secretStoreKey: "test-secret-store-key",
      localKmsMasterKey: "test-local-kms-master-key",
      awsKmsEnabled: false,
      retentionSweepIntervalMs: null
    };

    app = buildApp({ db, config });
    await app.ready();
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("rejects disallowed SIEM dataRegion when allow_cross_region_processing=false", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "siem-owner@example.com",
        password: "password1234",
        name: "Siem Owner",
        orgName: "SIEM Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const orgId = (register.json() as any).organization.id as string;

    const setPolicy = await app.inject({
      method: "PATCH",
      url: `/orgs/${orgId}/settings`,
      headers: { cookie },
      payload: {
        dataResidencyRegion: "eu",
        allowCrossRegionProcessing: false
      }
    });
    expect(setPolicy.statusCode).toBe(200);

    const disallowed = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      payload: {
        endpointUrl: "https://siem.example.com/ingest",
        dataRegion: "us"
      }
    });
    expect(disallowed.statusCode).toBe(400);
    expect((disallowed.json() as any).error).toBe("invalid_request");

    const allowed = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      payload: {
        endpointUrl: "https://siem.example.com/ingest",
        dataRegion: "eu"
      }
    });
    expect(allowed.statusCode).toBe(200);
    const body = allowed.json() as any;
    expect(body.config.endpointUrl).toBe("https://siem.example.com/ingest");
    expect(body.config.dataRegion).toBe("eu");

    const defaulted = await app.inject({
      method: "PUT",
      url: `/orgs/${orgId}/siem`,
      headers: { cookie },
      payload: {
        endpointUrl: "https://siem.example.com/ingest"
      }
    });
    expect(defaulted.statusCode).toBe(200);
    expect((defaulted.json() as any).config.dataRegion).toBe("eu");
  });
});

