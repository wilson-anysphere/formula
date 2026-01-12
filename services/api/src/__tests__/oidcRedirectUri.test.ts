import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
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

async function startDiscoveryServer(): Promise<{ issuerUrl: string; close: () => Promise<void> }> {
  const server = http.createServer((req, res) => {
    const baseUrl = `http://${req.headers.host}`;
    const url = new URL(req.url ?? "/", baseUrl);

    if (url.pathname === "/.well-known/openid-configuration") {
      const issuer = baseUrl;
      res.writeHead(200, { "content-type": "application/json" });
      res.end(
        JSON.stringify({
          issuer,
          authorization_endpoint: `${issuer}/authorize`,
          token_endpoint: `${issuer}/token`,
          jwks_uri: `${issuer}/jwks`
        })
      );
      return;
    }

    res.writeHead(404, { "content-type": "text/plain" });
    res.end("not found");
  });

  await new Promise<void>((resolve) => {
    server.listen(0, "127.0.0.1", () => resolve());
  });
  const addr = server.address();
  if (!addr || typeof addr === "string") throw new Error("expected discovery server to listen on tcp port");

  const issuerUrl = `http://127.0.0.1:${addr.port}`;
  return {
    issuerUrl,
    close: async () => {
      await new Promise<void>((resolve, reject) => {
        server.close((err) => (err ? reject(err) : resolve()));
      });
    }
  };
}

describe("OIDC redirect_uri construction", () => {
  let db: Pool;
  let config: AppConfig;
  let app: ReturnType<typeof buildApp>;
  let issuerUrl: string;
  let closeIssuer: () => Promise<void>;

  beforeAll(async () => {
    const issuer = await startDiscoveryServer();
    issuerUrl = issuer.issuerUrl;
    closeIssuer = issuer.close;

    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

    config = {
      port: 0,
      databaseUrl: "postgres://unused",
      publicBaseUrl: "https://trusted.example.com",
      publicBaseUrlHostAllowlist: ["trusted.example.com"],
      trustProxy: true,
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
    await closeIssuer();
  });

  it("ignores spoofed forwarded host headers when PUBLIC_BASE_URL is configured", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "oidc-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Redirect Test Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const orgId = (register.json() as any).organization.id as string;

    await db.query("UPDATE org_settings SET allowed_auth_methods = $2::jsonb WHERE org_id = $1", [
      orgId,
      JSON.stringify(["password", "oidc"])
    ]);

    await db.query(
      `
        INSERT INTO org_oidc_providers (org_id, provider_id, issuer_url, client_id, scopes, enabled)
        VALUES ($1,$2,$3,$4,$5::jsonb,$6)
      `,
      [orgId, "mock", issuerUrl, "client", JSON.stringify(["openid", "email"]), true]
    );

    await putSecret(db, config.secretStoreKeys, `oidc:${orgId}:mock`, "test-client-secret");

    const startRes = await app.inject({
      method: "GET",
      url: `/auth/oidc/${orgId}/mock/start`,
      headers: {
        // If redirect URI construction were derived from request headers, this would
        // taint redirect_uri.
        "x-forwarded-host": "evil.example.com",
        "x-forwarded-proto": "https"
      }
    });
    expect(startRes.statusCode).toBe(302);
    const location = startRes.headers.location as string;
    const authUrl = new URL(location);
    const redirectUri = authUrl.searchParams.get("redirect_uri");
    expect(redirectUri).toBe(`https://trusted.example.com/auth/oidc/${orgId}/mock/callback`);
  });
});
