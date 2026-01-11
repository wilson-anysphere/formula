import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { buildApp } from "../../../../services/api/src/app";
import { runMigrations } from "../../../../services/api/src/db/migrations";
import { deriveSecretStoreKey } from "../../../../services/api/src/secrets/secretStore";
import { ApiVersionStore } from "./apiVersionStore.js";
import { VersionManager } from "../versioning/versionManager.js";
import { EventEmitter } from "node:events";

function getMigrationsDir(): string {
  const here = path.dirname(fileURLToPath(import.meta.url));
  // packages/versioning/src/store -> services/api/migrations
  return path.resolve(here, "../../../../services/api/migrations");
}

function extractCookieFromFetch(response: Response): string {
  // undici / Node fetch exposes `getSetCookie()`, but fall back to the raw header.
  const headers = response.headers as any;
  const setCookies: string[] =
    typeof headers.getSetCookie === "function"
      ? (headers.getSetCookie() as string[])
      : [response.headers.get("set-cookie")].filter(Boolean) as string[];

  if (!setCookies.length) throw new Error("missing set-cookie header");
  return setCookies[0].split(";")[0];
}

class FakeDoc extends EventEmitter {
  private state: Uint8Array;

  constructor(initial: Uint8Array) {
    super();
    this.state = initial;
  }

  encodeState(): Uint8Array {
    return this.state;
  }

  applyState(snapshot: Uint8Array): void {
    this.state = new Uint8Array(snapshot);
    this.emit("update");
  }

  setState(next: Uint8Array): void {
    this.state = next;
    this.emit("update");
  }
}

describe("ApiVersionStore (integration): persists VersionManager history via services/api", () => {
  let db: any;
  let app: any;
  let baseUrl: string;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

    const config = {
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

    app = buildApp({ db, config });
    await app.ready();
    await app.listen({ port: 0, host: "127.0.0.1" });
    const address = app.server.address();
    if (!address || typeof address === "string") throw new Error("failed to start api server");
    baseUrl = `http://127.0.0.1:${address.port}`;
  });

  afterAll(async () => {
    await app?.close?.();
    await db?.end?.();
  });

  it("round-trips snapshots/checkpoints, lists, fetches, and deletes via cloud API", async () => {
    const registerRes = await fetch(`${baseUrl}/auth/register`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        email: "versions@example.com",
        password: "password1234",
        name: "Versions",
        orgName: "Version Org"
      })
    });
    expect(registerRes.status).toBe(200);
    const cookie = extractCookieFromFetch(registerRes);
    const registerBody = (await registerRes.json()) as any;
    const orgId = registerBody.organization.id as string;
    const userId = registerBody.user.id as string;

    const createDocRes = await fetch(`${baseUrl}/docs`, {
      method: "POST",
      headers: { "content-type": "application/json", cookie },
      body: JSON.stringify({ orgId, title: "Cloud versions doc" })
    });
    expect(createDocRes.status).toBe(200);
    const docId = ((await createDocRes.json()) as any).document.id as string;

    const store = new ApiVersionStore({
      baseUrl,
      docId,
      auth: { cookie }
    });

    const doc = new FakeDoc(new Uint8Array([1, 2, 3]));
    const manager = new VersionManager({
      doc,
      store,
      user: { userId, userName: "Versions" },
      autoStart: false
    });

    const checkpointBytes = new Uint8Array([10, 11, 12, 13]);
    doc.setState(checkpointBytes);
    const checkpoint = await manager.createCheckpoint({
      name: "Milestone",
      annotations: "ship it",
      locked: false
    });

    const snapshotBytes = new Uint8Array([99, 98, 97]);
    doc.setState(snapshotBytes);
    const snapshot = await manager.createSnapshot({ description: "Autosave" });

    const versions = await manager.listVersions();
    expect(versions.map((v) => v.id).sort()).toEqual([checkpoint.id, snapshot.id].sort());

    const roundTripCheckpoint = await manager.getVersion(checkpoint.id);
    expect(roundTripCheckpoint?.kind).toBe("checkpoint");
    expect(roundTripCheckpoint?.checkpointName).toBe("Milestone");
    expect(roundTripCheckpoint?.checkpointAnnotations).toBe("ship it");
    expect(roundTripCheckpoint?.checkpointLocked).toBe(false);
    expect(Array.from(roundTripCheckpoint!.snapshot)).toEqual(Array.from(checkpointBytes));

    const roundTripSnapshot = await manager.getVersion(snapshot.id);
    expect(roundTripSnapshot?.kind).toBe("snapshot");
    expect(roundTripSnapshot?.description).toBe("Autosave");
    expect(Array.from(roundTripSnapshot!.snapshot)).toEqual(Array.from(snapshotBytes));

    await manager.deleteVersion(snapshot.id);
    const afterDelete = await manager.listVersions();
    expect(afterDelete.map((v) => v.id)).toEqual([checkpoint.id]);
  });
});
