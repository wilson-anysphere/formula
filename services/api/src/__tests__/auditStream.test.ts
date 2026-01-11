import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import crypto from "node:crypto";
import type { AppConfig } from "../config";
import { buildApp } from "../app";
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

type ParsedSseEvent = { id: string; data: any };

class TestSseClient {
  private buffer = "";
  private readonly decoder = new TextDecoder();

  private constructor(
    private readonly reader: ReadableStreamDefaultReader<Uint8Array>,
    private readonly controller: AbortController
  ) {}

  static async connect(url: string, options: { headers: Record<string, string> }): Promise<TestSseClient> {
    const controller = new AbortController();
    const res = await fetch(url, {
      headers: { accept: "text/event-stream", ...options.headers },
      signal: controller.signal
    });
    if (res.status !== 200) {
      controller.abort();
      throw new Error(`expected 200, got ${res.status}`);
    }
    if (!res.body) {
      controller.abort();
      throw new Error("missing response body");
    }
    return new TestSseClient(res.body.getReader(), controller);
  }

  async nextEvent(options: { timeoutMs?: number } = {}): Promise<ParsedSseEvent> {
    const timeoutMs = options.timeoutMs ?? 5000;

    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const parsed = this.tryParseEvent();
      if (parsed) return parsed;

      const remaining = deadline - Date.now();
      const readPromise = this.reader.read();
      const result = await Promise.race([
        readPromise,
        new Promise<null>((resolve) => setTimeout(() => resolve(null), Math.max(0, remaining)))
      ]);
      if (result === null) break;

      if (result.done) throw new Error("SSE stream closed before event received");
      this.buffer += this.decoder.decode(result.value, { stream: true });
      this.buffer = this.buffer.replaceAll("\r\n", "\n");
    }

    throw new Error("timed out waiting for SSE event");
  }

  close(): void {
    this.controller.abort();
    void this.reader.cancel().catch(() => {});
  }

  private tryParseEvent(): ParsedSseEvent | null {
    const endIdx = this.buffer.indexOf("\n\n");
    if (endIdx === -1) return null;
    const frame = this.buffer.slice(0, endIdx);
    this.buffer = this.buffer.slice(endIdx + 2);

    let id: string | null = null;
    const dataLines: string[] = [];

    for (const line of frame.split("\n")) {
      if (line.startsWith(":")) continue;
      if (line.startsWith("id:")) {
        id = line.slice("id:".length).trim();
        continue;
      }
      if (line.startsWith("data:")) {
        dataLines.push(line.slice("data:".length).trimStart());
      }
    }

    if (!id || dataLines.length === 0) return null;

    const rawData = dataLines.join("\n");
    return { id, data: JSON.parse(rawData) };
  }
}

describe("audit ingest + SSE streaming", () => {
  let db: Pool;
  let config: AppConfig;
  let app: ReturnType<typeof buildApp>;
  let baseUrl: string;

  let orgId: string;
  let adminCookie: string;
  let adminUserId: string;
  let memberCookie: string;

  beforeAll(async () => {
    const mem = newDb({ autoCreateForeignKeyIndices: true });
    const pgAdapter = mem.adapters.createPg();
    db = new pgAdapter.Pool();
    await runMigrations(db, { migrationsDir: getMigrationsDir() });

    config = {
      port: 0,
      databaseUrl: "postgres://unused",
      sessionCookieName: "formula_session",
      sessionTtlSeconds: 60 * 60,
      cookieSecure: false,
      publicBaseUrlHostAllowlist: ["localhost", "127.0.0.1"],
      corsAllowedOrigins: [],
      syncTokenSecret: "test-sync-secret",
      syncTokenTtlSeconds: 60,
      secretStoreKeys: {
        currentKeyId: "test",
        keys: { test: deriveSecretStoreKey("test-secret-store-key") }
      },
      localKmsMasterKey: "test-local-kms-master-key",
      awsKmsEnabled: false,
      retentionSweepIntervalMs: null,
      oidcAuthStateCleanupIntervalMs: null
    };

    app = buildApp({ db, config });
    await app.ready();
    await app.listen({ port: 0, host: "127.0.0.1" });
    const addr = app.server.address();
    if (!addr || typeof addr === "string") throw new Error("expected server to listen on a tcp port");
    baseUrl = `http://127.0.0.1:${addr.port}`;

    const registerAdmin = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "admin@example.com",
        password: "password1234",
        name: "Admin",
        orgName: "Acme"
      }
    });
    expect(registerAdmin.statusCode).toBe(200);
    adminCookie = extractCookie(registerAdmin.headers["set-cookie"]);
    const adminBody = registerAdmin.json() as any;
    orgId = adminBody.organization.id;
    adminUserId = adminBody.user.id;

    const registerMember = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "member@example.com",
        password: "password1234",
        name: "Member"
      }
    });
    expect(registerMember.statusCode).toBe(200);
    memberCookie = extractCookie(registerMember.headers["set-cookie"]);
    const memberUserId = (registerMember.json() as any).user.id as string;

    await db.query("INSERT INTO org_members (org_id, user_id, role) VALUES ($1, $2, 'member')", [
      orgId,
      memberUserId
    ]);
  }, 30_000);

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("POST /orgs/:orgId/audit writes an event and it appears in GET", async () => {
    const ingest = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/audit`,
      headers: { cookie: adminCookie },
      payload: {
        eventType: "test.ingest",
        resource: { type: "organization", id: orgId },
        success: true,
        details: { hello: "world" }
      }
    });
    expect(ingest.statusCode).toBe(202);
    const ingestId = (ingest.json() as any).id as string;
    expect(ingestId).toMatch(/^[0-9a-f-]{36}$/i);

    const list = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/audit?limit=50`,
      headers: { cookie: adminCookie }
    });
    expect(list.statusCode).toBe(200);
    const events = (list.json() as any).events as any[];
    expect(events.some((e) => e.id === ingestId && e.eventType === "test.ingest")).toBe(true);
  });

  it("SSE stream requires org admin", async () => {
    const res = await fetch(`${baseUrl}/orgs/${orgId}/audit/stream`, { headers: { cookie: memberCookie } });
    expect(res.status).toBe(403);
  });

  it("SSE streams events written by POST /orgs/:orgId/audit (same process)", async () => {
    const client = await TestSseClient.connect(`${baseUrl}/orgs/${orgId}/audit/stream`, {
      headers: { cookie: adminCookie }
    });

    try {
      const ingest = await app.inject({
        method: "POST",
        url: `/orgs/${orgId}/audit`,
        headers: { cookie: adminCookie },
        payload: {
          eventType: "test.stream_local",
          resource: { type: "organization", id: orgId },
          success: true,
          details: { foo: "bar" }
        }
      });
      expect(ingest.statusCode).toBe(202);
      const ingestId = (ingest.json() as any).id as string;

      const evt = await client.nextEvent();
      expect(evt.data.id).toBe(ingestId);
      expect(evt.data.eventType).toBe("test.stream_local");
    } finally {
      client.close();
    }
  });

  it("SSE streams events delivered via pg NOTIFY fanout (simulated)", async () => {
    const client = await TestSseClient.connect(`${baseUrl}/orgs/${orgId}/audit/stream`, {
      headers: { cookie: adminCookie }
    });

    const id = crypto.randomUUID();
    await db.query(
      `
        INSERT INTO audit_log (id, org_id, user_id, event_type, resource_type, success, details, created_at)
        VALUES ($1, $2, $3, $4, 'organization', true, $5::jsonb, now())
      `,
      [id, orgId, adminUserId, "test.stream_notify", JSON.stringify({ token: "super-secret" })]
    );

    await app.auditStreamHub.injectNotification({ orgId, id });

    try {
      const evt = await client.nextEvent();
      expect(evt.data.id).toBe(id);
      expect(evt.data.eventType).toBe("test.stream_notify");
      expect(evt.data.details).toMatchObject({ token: "[REDACTED]" });
    } finally {
      client.close();
    }
  });

  it("SSE supports Last-Event-ID resume", async () => {
    const firstId = crypto.randomUUID();
    const secondId = crypto.randomUUID();
    const t0 = new Date("2099-01-01T00:00:00.000Z");
    const t1 = new Date("2099-01-01T00:00:01.000Z");

    await db.query(
      `
        INSERT INTO audit_log (id, org_id, user_id, event_type, resource_type, success, details, created_at)
        VALUES ($1, $2, $3, $4, 'organization', true, '{}'::jsonb, $5)
      `,
      [firstId, orgId, adminUserId, "test.resume.first", t0]
    );
    await db.query(
      `
        INSERT INTO audit_log (id, org_id, user_id, event_type, resource_type, success, details, created_at)
        VALUES ($1, $2, $3, $4, 'organization', true, '{}'::jsonb, $5)
      `,
      [secondId, orgId, adminUserId, "test.resume.second", t1]
    );

    const client = await TestSseClient.connect(`${baseUrl}/orgs/${orgId}/audit/stream`, {
      headers: { cookie: adminCookie, "last-event-id": firstId }
    });

    try {
      const evt = await client.nextEvent();
      expect(evt.data.id).toBe(secondId);
      expect(evt.data.eventType).toBe("test.resume.second");
    } finally {
      client.close();
    }
  });
});
