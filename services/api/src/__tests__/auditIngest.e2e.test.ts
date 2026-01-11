import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { newDb } from "pg-mem";
import type { Pool } from "pg";
import path from "node:path";
import { fileURLToPath } from "node:url";
import http from "node:http";
import type { Readable } from "node:stream";
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

type SseMessage = {
  id?: string;
  event?: string;
  data?: unknown;
};

async function waitForSse(
  stream: Readable,
  predicate: (msg: SseMessage) => boolean,
  timeoutMs = 10_000
): Promise<SseMessage> {
  return await new Promise<SseMessage>((resolve, reject) => {
    let buffer = "";

    const onData = (chunk: Buffer | string) => {
      buffer += String(chunk).replaceAll("\r\n", "\n");

      while (true) {
        const idx = buffer.indexOf("\n\n");
        if (idx === -1) break;

        const frame = buffer.slice(0, idx);
        buffer = buffer.slice(idx + 2);

        // Comments (keep-alives) start with ":".
        if (frame.trim().startsWith(":")) continue;

        const msg: SseMessage = {};
        const dataLines: string[] = [];
        for (const line of frame.split("\n")) {
          if (line.startsWith("id:")) msg.id = line.slice("id:".length).trim();
          if (line.startsWith("event:")) msg.event = line.slice("event:".length).trim();
          if (line.startsWith("data:")) dataLines.push(line.slice("data:".length).trimStart());
        }

        if (dataLines.length > 0) {
          const joined = dataLines.join("\n");
          try {
            msg.data = JSON.parse(joined);
          } catch {
            msg.data = joined;
          }
        }

        if (predicate(msg)) {
          cleanup();
          resolve(msg);
          return;
        }
      }
    };

    const onError = (err: unknown) => {
      cleanup();
      reject(err);
    };

    const onEnd = () => {
      cleanup();
      reject(new Error("SSE stream ended before expected message"));
    };

    const timeout = setTimeout(() => {
      cleanup();
      reject(new Error("Timed out waiting for SSE message"));
    }, timeoutMs);
    timeout.unref?.();

    const cleanup = () => {
      clearTimeout(timeout);
      stream.off("data", onData);
      stream.off("error", onError);
      stream.off("end", onEnd);
    };

    stream.on("data", onData);
    stream.on("error", onError);
    stream.on("end", onEnd);
  });
}

describe("API e2e: audit ingestion + streaming", () => {
  let db: Pool;
  let config: AppConfig;
  let app: ReturnType<typeof buildApp>;
  let baseUrl: URL;

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

    const address = await app.listen({ port: 0, host: "127.0.0.1" });
    baseUrl = new URL(address);
  });

  afterAll(async () => {
    await app.close();
    await db.end();
  });

  it("ingests audit events and stores server-derived actor/context", async () => {
    const register = await app.inject({
      method: "POST",
      url: "/auth/register",
      payload: {
        email: "audit-owner@example.com",
        password: "password1234",
        name: "Owner",
        orgName: "Audit Org"
      }
    });
    expect(register.statusCode).toBe(200);
    const cookie = extractCookie(register.headers["set-cookie"]);
    const registerBody = register.json() as any;
    const orgId = registerBody.organization.id as string;
    const userId = registerBody.user.id as string;

    const ingest = await app.inject({
      method: "POST",
      url: `/orgs/${orgId}/audit`,
      remoteAddress: "203.0.113.5",
      headers: {
        cookie,
        "user-agent": "UnitTest/1.0"
      },
      payload: {
        eventType: "client.button_clicked",
        resource: { type: "document", id: "doc_123", name: "Q1 Plan" },
        success: true,
        details: { token: "super-secret", foo: "bar" }
      }
    });

    expect(ingest.statusCode).toBe(202);
    const eventId = (ingest.json() as any).id as string;
    expect(eventId).toBeTypeOf("string");

    const audit = await app.inject({
      method: "GET",
      url: `/orgs/${orgId}/audit?eventType=client.button_clicked`,
      headers: { cookie }
    });
    expect(audit.statusCode).toBe(200);
    const events = (audit.json() as any).events as any[];
    const found = events.find((event) => event.id === eventId);
    expect(found).toBeTruthy();
    expect(found).toMatchObject({
      id: eventId,
      eventType: "client.button_clicked",
      actor: { type: "user", id: userId },
      context: {
        orgId,
        userId,
        userEmail: "audit-owner@example.com",
        ipAddress: "203.0.113.5",
        userAgent: "UnitTest/1.0"
      },
      resource: {
        type: "document",
        id: "doc_123",
        name: "Q1 Plan"
      },
      success: true,
      details: { foo: "bar", token: "[REDACTED]" }
    });
    // Internal storage metadata must never be exposed to API consumers.
    expect(found.details).not.toHaveProperty("__audit");
    // Server must assign a canonical timestamp/id.
    expect(new Date(found.timestamp).toISOString()).toBe(found.timestamp);

    const stored = await db.query(
      "SELECT org_id, user_id, user_email, ip_address, user_agent, session_id, created_at FROM audit_log WHERE id = $1",
      [eventId]
    );
    expect(stored.rowCount).toBe(1);
    expect(stored.rows[0]!.org_id).toBe(orgId);
    expect(stored.rows[0]!.user_id).toBe(userId);
    expect(stored.rows[0]!.user_email).toBe("audit-owner@example.com");
    expect(stored.rows[0]!.ip_address).toBe("203.0.113.5");
    expect(stored.rows[0]!.user_agent).toBe("UnitTest/1.0");
    expect(stored.rows[0]!.session_id).toBeTruthy();
    expect(stored.rows[0]!.created_at).toBeTruthy();
  });

  it(
    "streams audit events over SSE (admin-only) with redaction",
    async () => {
      const register = await app.inject({
        method: "POST",
        url: "/auth/register",
        payload: {
          email: "audit-stream-owner@example.com",
          password: "password1234",
          name: "Owner",
          orgName: "Stream Org"
        }
      });
      expect(register.statusCode).toBe(200);
      const cookie = extractCookie(register.headers["set-cookie"]);
      const registerBody = register.json() as any;
      const orgId = registerBody.organization.id as string;

      const res = await new Promise<http.IncomingMessage>((resolve, reject) => {
        const req = http.request(
          {
            method: "GET",
            hostname: baseUrl.hostname,
            port: Number(baseUrl.port),
            path: `/orgs/${orgId}/audit/stream`,
            headers: {
              cookie
            }
          },
          resolve
        );
        req.on("error", reject);
        req.end();
      });
      expect(res.statusCode).toBe(200);
      res.setEncoding("utf8");

      const waitForEvent = waitForSse(res as unknown as Readable, (msg) => {
        return (msg.data as any)?.eventType === "client.stream_test" && msg.event === "audit";
      });

      const ingest = await app.inject({
        method: "POST",
        url: `/orgs/${orgId}/audit`,
        remoteAddress: "203.0.113.9",
        headers: {
          cookie,
          "user-agent": "UnitTest/2.0"
        },
        payload: {
          eventType: "client.stream_test",
          details: {
            token: "stream-secret",
            nested: { secret: "nested-secret" }
          }
        }
      });
      expect(ingest.statusCode).toBe(202);

      const msg = await waitForEvent;
      expect(msg.event).toBe("audit");
      const event = msg.data as any;
      expect(event).toMatchObject({
        eventType: "client.stream_test",
        context: {
          orgId,
          ipAddress: "203.0.113.9",
          userAgent: "UnitTest/2.0"
        },
        details: {
          token: "[REDACTED]",
          nested: { secret: "[REDACTED]" }
        }
      });
      // Internal storage metadata must never be exposed over SSE.
      expect(event.details).not.toHaveProperty("__audit");

      res.destroy();
    },
    20_000
  );
});
