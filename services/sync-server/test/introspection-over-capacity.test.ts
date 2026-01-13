import assert from "node:assert/strict";
import http from "node:http";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";

import { getAvailablePort, startSyncServer } from "./test-helpers.ts";

const JWT_SECRET = "test-secret";
const JWT_AUDIENCE = "formula-sync";

function signJwt(payload: Record<string, unknown>): string {
  return jwt.sign(payload, JWT_SECRET, {
    algorithm: "HS256",
    audience: JWT_AUDIENCE,
  });
}

type WsUpgradeOutcome =
  | { outcome: "open" }
  | { outcome: "rejected"; statusCode: number | undefined }
  | { outcome: "error"; error: unknown };

function attemptWebSocketUpgrade(url: string): Promise<WsUpgradeOutcome> {
  return new Promise((resolve) => {
    const ws = new WebSocket(url);
    let finished = false;

    const finish = (result: WsUpgradeOutcome) => {
      if (finished) return;
      finished = true;
      try {
        ws.terminate();
      } catch {
        // ignore
      }
      resolve(result);
    };

    ws.on("open", () => finish({ outcome: "open" }));
    ws.on("unexpected-response", (_req, res) => {
      const statusCode = res.statusCode;
      res.resume();
      finish({ outcome: "rejected", statusCode });
    });
    ws.on("error", (err) => finish({ outcome: "error", error: err }));
  });
}

function parseCounterValue(body: string, name: string, labels?: Record<string, string>): number {
  const escapeRegExp = (value: string): string =>
    value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  const escapedName = escapeRegExp(name);

  if (!labels || Object.keys(labels).length === 0) {
    const m = body.match(
      new RegExp(`^${escapedName}(?:\\{[^}]*\\})?\\s+(\\d+(?:\\.\\d+)?)$`, "m")
    );
    return m ? Number.parseFloat(m[1]!) : 0;
  }

  const lookaheads = Object.entries(labels)
    .map(([k, v]) => `(?=[^}]*${escapeRegExp(`${k}="${v}"`)})`)
    .join("");
  const re = new RegExp(
    `^${escapedName}\\{${lookaheads}[^}]*\\}\\s+(\\d+(?:\\.\\d+)?)$`,
    "m"
  );
  const m = body.match(re);
  return m ? Number.parseFloat(m[1]!) : 0;
}

test("sync-server fails fast with 503 when sync token introspection is over capacity", async (t) => {
  const docName = `doc-${Math.random().toString(16).slice(2)}`;

  const tokenA = signJwt({
    sub: "u-a",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-a",
  });
  const tokenB = signJwt({
    sub: "u-b",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-b",
  });
  const tokenC = signJwt({
    sub: "u-c",
    docId: docName,
    orgId: "o1",
    role: "editor",
    sessionId: "s-c",
  });

  const introspectionAdminToken = "introspection-admin-token";
  const introspectionPort = await getAvailablePort();
  const calls: string[] = [];

  const introspectionServer = http.createServer((req, res) => {
    void (async () => {
      if (req.method !== "POST" || req.url !== "/internal/sync/introspect") {
        res.writeHead(404).end();
        return;
      }

      const header = req.headers["x-internal-admin-token"];
      const provided =
        typeof header === "string" ? header : Array.isArray(header) ? header[0] : undefined;
      if (provided !== introspectionAdminToken) {
        res.writeHead(403, { "content-type": "application/json" });
        res.end(JSON.stringify({ ok: false, active: false, error: "forbidden" }));
        return;
      }

      const chunks: Buffer[] = [];
      for await (const chunk of req) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      }
      const parsed = JSON.parse(Buffer.concat(chunks).toString("utf8")) as any;
      calls.push(String(parsed.token ?? ""));

      // Hold the first request long enough for other websocket upgrade attempts to
      // hit the concurrency limit.
      await new Promise((resolve) => setTimeout(resolve, 500));

      res.writeHead(200, { "content-type": "application/json" });
      res.end(JSON.stringify({ active: true, userId: "u-ok", orgId: "o1", role: "editor" }));
    })().catch((err) => {
      res.writeHead(500, { "content-type": "application/json" });
      res.end(JSON.stringify({ error: "internal_error", message: String(err) }));
    });
  });

  await new Promise<void>((resolve) => {
    introspectionServer.listen(introspectionPort, "127.0.0.1", () => resolve());
  });
  t.after(async () => {
    await new Promise<void>((resolve, reject) =>
      introspectionServer.close((err) => (err ? reject(err) : resolve()))
    );
  });

  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-introspection-over-capacity-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const syncServer = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: {
      SYNC_SERVER_INTROSPECTION_URL: `http://127.0.0.1:${introspectionPort}`,
      SYNC_SERVER_INTROSPECTION_TOKEN: introspectionAdminToken,
      SYNC_SERVER_INTROSPECTION_CACHE_TTL_MS: "0",
      SYNC_SERVER_INTROSPECTION_MAX_CONCURRENT: "1",
    },
  });
  t.after(async () => {
    await syncServer.stop();
  });

  const [a, b, c] = await Promise.all([
    attemptWebSocketUpgrade(`${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(tokenA)}`),
    attemptWebSocketUpgrade(`${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(tokenB)}`),
    attemptWebSocketUpgrade(`${syncServer.wsUrl}/${docName}?token=${encodeURIComponent(tokenC)}`),
  ]);

  const results = [a, b, c];
  const opens = results.filter((r) => r.outcome === "open").length;
  const rejected503 = results.filter(
    (r) => r.outcome === "rejected" && r.statusCode === 503
  ).length;
  const errors = results.filter((r) => r.outcome === "error");

  assert.equal(errors.length, 0, `Unexpected websocket errors: ${JSON.stringify(errors)}`);
  assert.ok(opens >= 1, `Expected at least one successful upgrade, got: ${JSON.stringify(results)}`);
  assert.ok(
    rejected503 >= 1,
    `Expected at least one 503 rejection, got: ${JSON.stringify(results)}`
  );

  // Only one request should make it through to the (slow) introspection server.
  assert.equal(calls.length, 1);
  assert.ok(calls[0] === tokenA || calls[0] === tokenB || calls[0] === tokenC);

  const metricsRes = await fetch(`${syncServer.httpUrl}/metrics`);
  assert.equal(metricsRes.status, 200);
  const metricsBody = await metricsRes.text();

  assert.match(metricsBody, /sync_server_introspection_over_capacity_total/);
  assert.match(metricsBody, /sync_server_introspection_requests_total/);

  const overCapacity = parseCounterValue(
    metricsBody,
    "sync_server_introspection_over_capacity_total"
  );
  const okCount = parseCounterValue(metricsBody, "sync_server_introspection_requests_total", {
    result: "ok",
  });
  const errorCount = parseCounterValue(metricsBody, "sync_server_introspection_requests_total", {
    result: "error",
  });

  assert.ok(overCapacity >= 1, `Expected over capacity >= 1, got ${overCapacity}`);
  assert.ok(okCount >= 1, `Expected ok >= 1, got ${okCount}`);
  assert.ok(errorCount >= 1, `Expected error >= 1, got ${errorCount}`);
});
