import assert from "node:assert/strict";
import https from "node:https";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import WebSocket from "ws";

import { startSyncServer } from "./test-helpers.ts";

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
  assert.match(body, /sync_server_persistence_info/);

  const health = await fetch(`${server.httpUrl}/healthz`);
  assert.equal(health.status, 200);
  assert.equal(health.headers.get("cache-control"), "no-store");

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
  assert.match(health.body, /\"status\":\"ok\"/);

  const stats = await httpsRequestText(`${server.httpUrl}/internal/stats`, {
    headers: { "x-internal-admin-token": "admin-token" },
  });
  assert.equal(stats.status, 200);
  assert.match(stats.body, /\"ok\":true/);

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
