import assert from "node:assert/strict";
import crypto from "node:crypto";
import http from "node:http";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";

import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";
import type { SyncServerConfig } from "../src/config.js";

const JWT_SECRET = "test-secret";
const JWT_AUDIENCE = "formula-sync";

function signJwt(payload: Record<string, unknown>): string {
  return jwt.sign(payload, JWT_SECRET, { algorithm: "HS256", audience: JWT_AUDIENCE });
}

function waitForWsOpen(ws: WebSocket): Promise<void> {
  return new Promise((resolve, reject) => {
    const onError = (err: unknown) => {
      ws.off("open", onOpen);
      reject(err);
    };
    const onOpen = () => {
      ws.off("error", onError);
      resolve();
    };
    ws.once("error", onError);
    ws.once("open", onOpen);
  });
}

function waitForWsClose(ws: WebSocket): Promise<void> {
  return new Promise((resolve) => {
    if (ws.readyState === WebSocket.CLOSED) return resolve();
    ws.once("close", () => resolve());
  });
}

test("optional JWT introspection cache is scoped per (token, docId, clientIp)", async (t) => {
  const docId = crypto.randomUUID();
  const token = signJwt({
    sub: "u1",
    docId,
    orgId: "o1",
    role: "editor",
    sessionId: crypto.randomUUID(),
  });

  const introspectionAdminToken = "introspection-admin-token";
  const hitsByKey = new Map<string, number>();

  const introspectionServer = http.createServer(async (req, res) => {
    if (req.method !== "POST" || req.url !== "/internal/sync/introspect") {
      res.writeHead(404).end();
      return;
    }

    const header = req.headers["x-internal-admin-token"];
    const provided =
      typeof header === "string"
        ? header
        : Array.isArray(header)
          ? header[0]
          : undefined;
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
    const bodyToken = parsed.token;
    const bodyDocId = parsed.docId;
    const clientIp = parsed.clientIp;

    if (bodyToken !== token || bodyDocId !== docId) {
      res.writeHead(400, { "content-type": "application/json" });
      res.end(JSON.stringify({ ok: false, active: false, error: "invalid_request" }));
      return;
    }

    assert.ok(typeof clientIp === "string" && clientIp.length > 0);

    const key = `${bodyToken}\n${bodyDocId}\n${clientIp}`;
    hitsByKey.set(key, (hitsByKey.get(key) ?? 0) + 1);

    res.writeHead(200, { "content-type": "application/json" });
    res.end(JSON.stringify({ active: true, userId: "u1", orgId: "o1", role: "editor" }));
  });

  await new Promise<void>((resolve) => introspectionServer.listen(0, "127.0.0.1", () => resolve()));
  t.after(async () => {
    await new Promise<void>((resolve, reject) =>
      introspectionServer.close((err) => (err ? reject(err) : resolve()))
    );
  });

  const addr = introspectionServer.address();
  assert.ok(addr && typeof addr !== "string");

  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-introspection-cache-ip-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const syncServerConfig: SyncServerConfig = {
    host: "127.0.0.1",
    port: 0,
    trustProxy: true,
    gc: true,
    shutdownGraceMs: 0,
    tls: null,
    metrics: { public: true },
    dataDir,
    disableDataDirLock: true,
    persistence: {
      backend: "file",
      compactAfterUpdates: 10,
      leveldbDocNameHashing: false,
      encryption: { mode: "off" },
    },
    auth: {
      mode: "jwt-hs256",
      secret: JWT_SECRET,
      audience: JWT_AUDIENCE,
      requireSub: false,
      requireExp: false,
    },
    enforceRangeRestrictions: false,
    introspection: {
      url: `http://127.0.0.1:${addr.port}/internal/sync/introspect`,
      token: introspectionAdminToken,
      cacheTtlMs: 30_000,
      maxConcurrent: 50,
    },
    internalAdminToken: null,
    retention: { ttlMs: 0, sweepIntervalMs: 0, tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000 },
    limits: {
      maxConnections: 100,
      maxConnectionsPerIp: 100,
      maxConnectionsPerDoc: 0,
      maxConnAttemptsPerWindow: 500,
      connAttemptWindowMs: 60_000,
      maxMessageBytes: 2 * 1024 * 1024,
      maxMessagesPerWindow: 5_000,
      messageWindowMs: 10_000,
      maxMessagesPerIpWindow: 0,
      ipMessageWindowMs: 0,
      maxAwarenessStateBytes: 64 * 1024,
      maxAwarenessEntries: 10,
      maxMessagesPerDocWindow: 10_000,
      docMessageWindowMs: 10_000,
      maxBranchingCommitsPerDoc: 0,
      maxVersionsPerDoc: 0,
    },
    logLevel: "silent",
  };

  const syncServer = createSyncServer(syncServerConfig, createLogger("silent"));
  await syncServer.start();
  t.after(async () => {
    await syncServer.stop();
  });

  const ipA = "1.2.3.4";
  const ipB = "5.6.7.8";

  const connectOnce = async (ip: string) => {
    const ws = new WebSocket(`${syncServer.getWsUrl()}/${docId}?token=${encodeURIComponent(token)}`, {
      headers: { "x-forwarded-for": ip },
    });
    try {
      await waitForWsOpen(ws);
      ws.close();
      await waitForWsClose(ws);
    } finally {
      try {
        ws.terminate();
      } catch {
        // ignore
      }
    }
  };

  await connectOnce(ipA);
  await connectOnce(ipA);
  await connectOnce(ipB);

  assert.equal(hitsByKey.get(`${token}\n${docId}\n${ipA}`), 1);
  assert.equal(hitsByKey.get(`${token}\n${docId}\n${ipB}`), 1);
});
