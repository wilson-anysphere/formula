import assert from "node:assert/strict";
import https from "node:https";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import WebSocket from "ws";

import { startSyncServer, waitForCondition } from "./test-helpers.ts";

function getMetricValue(body: string, name: string): number | null {
  // Allow default labels (e.g. `{service="sync-server"}`) while keeping the matcher simple.
  const match = body.match(
    new RegExp(
      `^${name}(?:\\{[^}]*\\})?\\s+(-?\\d+(?:\\.\\d+)?(?:e[+-]?\\d+)?)$`,
      "m"
    )
  );
  if (!match) return null;
  const value = Number(match[1]);
  return Number.isFinite(value) ? value : null;
}

async function httpsRequestText(url: string, opts?: { headers?: Record<string, string> }) {
  const parsed = new URL(url);
  assert.equal(parsed.protocol, "https:");

  return await new Promise<{ status: number; body: string }>((resolve, reject) => {
    const req = https.request(
      {
        protocol: parsed.protocol,
        hostname: parsed.hostname,
        port: parsed.port,
        path: `${parsed.pathname}${parsed.search}`,
        method: "GET",
        headers: opts?.headers,
        rejectUnauthorized: false,
      },
      (res) => {
        const chunks: Buffer[] = [];
        res.on("data", (d) => chunks.push(Buffer.from(d)));
        res.on("end", () => {
          resolve({
            status: res.statusCode ?? 0,
            body: Buffer.concat(chunks).toString("utf8"),
          });
        });
      }
    );
    req.on("error", reject);
    req.end();
  });
}

test("exposes Prometheus metrics in text format", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-metrics-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const ws = new WebSocket(`${server.wsUrl}/metrics-doc?token=test-token`);
  await new Promise<void>((resolve, reject) => {
    ws.once("open", () => resolve());
    ws.once("error", reject);
  });
  ws.terminate();

  const res = await fetch(`${server.httpUrl}/metrics`);
  assert.equal(res.status, 200);
  assert.match(res.headers.get("content-type") ?? "", /text\/plain/);
  assert.equal(res.headers.get("cache-control"), "no-store");
  const body = await res.text();
  assert.match(body, /sync_server_ws_connections_total/);
  assert.match(body, /sync_server_ws_closes_total/);
  assert.match(body, /sync_server_ws_messages_too_large_total/);
  assert.match(body, /sync_server_process_resident_memory_bytes/);
  assert.match(body, /sync_server_process_heap_used_bytes/);
  assert.match(body, /sync_server_process_heap_total_bytes/);
  assert.match(body, /sync_server_event_loop_delay_ms/);
  assert.match(body, /sync_server_shutdown_draining_current/);
  assert.match(body, /sync_server_persistence_info/);

  const health = await fetch(`${server.httpUrl}/healthz`);
  assert.equal(health.status, 200);
  assert.equal(health.headers.get("cache-control"), "no-store");
  const healthBody = (await health.json()) as {
    rssBytes?: unknown;
    heapUsedBytes?: unknown;
    heapTotalBytes?: unknown;
    eventLoopDelayMs?: unknown;
    draining?: unknown;
  };
  assert.equal(typeof healthBody.rssBytes, "number");
  assert.equal(typeof healthBody.heapUsedBytes, "number");
  assert.equal(typeof healthBody.heapTotalBytes, "number");
  assert.equal(typeof healthBody.eventLoopDelayMs, "number");
  assert.equal(typeof healthBody.draining, "boolean");

  const ready = await fetch(`${server.httpUrl}/readyz`);
  assert.equal(ready.status, 200);
  assert.equal(ready.headers.get("cache-control"), "no-store");

  const internalMissing = await fetch(`${server.httpUrl}/internal/metrics`);
  assert.equal(internalMissing.status, 403);

  const internalWrong = await fetch(`${server.httpUrl}/internal/metrics`, {
    headers: { "x-internal-admin-token": "wrong-token" },
  });
  assert.equal(internalWrong.status, 403);

  const internalOk = await fetch(`${server.httpUrl}/internal/metrics`, {
    headers: { "x-internal-admin-token": "admin-token" },
  });
  assert.equal(internalOk.status, 200);
  assert.match(internalOk.headers.get("content-type") ?? "", /text\/plain/);
  assert.equal(internalOk.headers.get("cache-control"), "no-store");
  const internalBody = await internalOk.text();
  assert.match(internalBody, /sync_server_ws_connections_total/);
});

test("updates active doc + unique IP gauges as websocket connections come and go", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-metrics-gauges-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const ws1 = new WebSocket(`${server.wsUrl}/doc-a?token=test-token`);
  const ws2 = new WebSocket(`${server.wsUrl}/doc-b?token=test-token`);

  await Promise.all(
    [ws1, ws2].map(
      (ws) =>
        new Promise<void>((resolve, reject) => {
          ws.once("open", () => resolve());
          ws.once("error", reject);
        })
    )
  );

  await waitForCondition(async () => {
    const res = await fetch(`${server.httpUrl}/metrics`);
    if (res.status !== 200) return false;
    const body = await res.text();
    const activeDocs = getMetricValue(body, "sync_server_ws_active_docs_current");
    const uniqueIps = getMetricValue(body, "sync_server_ws_unique_ips_current");
    return activeDocs === 2 && uniqueIps !== null && uniqueIps >= 1;
  }, 5_000);

  const closePromises = [ws1, ws2].map(
    (ws) =>
      new Promise<void>((resolve) => {
        ws.once("close", () => resolve());
        ws.once("error", () => resolve());
      })
  );
  ws1.terminate();
  ws2.terminate();
  await Promise.all(closePromises);

  await waitForCondition(async () => {
    const res = await fetch(`${server.httpUrl}/metrics`);
    if (res.status !== 200) return false;
    const body = await res.text();
    const activeDocs = getMetricValue(body, "sync_server_ws_active_docs_current");
    const uniqueIps = getMetricValue(body, "sync_server_ws_unique_ips_current");
    return activeDocs === 0 && uniqueIps === 0;
  }, 5_000);
});

test("supports HTTPS/WSS when SYNC_SERVER_TLS_CERT_PATH and SYNC_SERVER_TLS_KEY_PATH are set", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-tls-"));

  const fixturesDir = path.join(path.dirname(fileURLToPath(import.meta.url)), "fixtures");
  const certPath = path.join(fixturesDir, "localhost-cert.pem");
  const keyPath = path.join(fixturesDir, "localhost-key.pem");

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
      SYNC_SERVER_TLS_CERT_PATH: certPath,
      SYNC_SERVER_TLS_KEY_PATH: keyPath,
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  assert.match(server.httpUrl, /^https:\/\//);
  assert.match(server.wsUrl, /^wss:\/\//);

  const health = await httpsRequestText(`${server.httpUrl}/healthz`);
  assert.equal(health.status, 200);
  const healthJson = JSON.parse(health.body) as {
    status?: unknown;
    rssBytes?: unknown;
    heapUsedBytes?: unknown;
    heapTotalBytes?: unknown;
    eventLoopDelayMs?: unknown;
  };
  assert.equal(healthJson.status, "ok");
  assert.equal(typeof healthJson.rssBytes, "number");
  assert.equal(typeof healthJson.heapUsedBytes, "number");
  assert.equal(typeof healthJson.heapTotalBytes, "number");
  assert.equal(typeof healthJson.eventLoopDelayMs, "number");

  const stats = await httpsRequestText(`${server.httpUrl}/internal/stats`, {
    headers: { "x-internal-admin-token": "admin-token" },
  });
  assert.equal(stats.status, 200);
  const statsJson = JSON.parse(stats.body) as {
    ok?: unknown;
    rssBytes?: unknown;
    heapUsedBytes?: unknown;
    heapTotalBytes?: unknown;
    eventLoopDelayMs?: unknown;
    connections?: { activeDocs?: unknown };
  };
  assert.equal(statsJson.ok, true);
  assert.equal(typeof statsJson.rssBytes, "number");
  assert.equal(typeof statsJson.heapUsedBytes, "number");
  assert.equal(typeof statsJson.heapTotalBytes, "number");
  assert.equal(typeof statsJson.eventLoopDelayMs, "number");
  assert.equal(typeof statsJson.connections?.activeDocs, "number");

  const ws = new WebSocket(`${server.wsUrl}/tls-doc?token=test-token`, {
    rejectUnauthorized: false,
  });
  await new Promise<void>((resolve, reject) => {
    ws.once("open", () => resolve());
    ws.once("error", reject);
  });
  ws.terminate();
});

test("can disable public /metrics while keeping /internal/metrics", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-metrics-disabled-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_PERSISTENCE_ENCRYPTION: "off",
      SYNC_SERVER_INTERNAL_ADMIN_TOKEN: "admin-token",
      SYNC_SERVER_DISABLE_PUBLIC_METRICS: "true",
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const publicRes = await fetch(`${server.httpUrl}/metrics`);
  assert.equal(publicRes.status, 404);

  const internalRes = await fetch(`${server.httpUrl}/internal/metrics`, {
    headers: { "x-internal-admin-token": "admin-token" },
  });
  assert.equal(internalRes.status, 200);
  assert.match(internalRes.headers.get("content-type") ?? "", /text\/plain/);
  const body = await internalRes.text();
  assert.match(body, /sync_server_ws_connections_total/);
});
