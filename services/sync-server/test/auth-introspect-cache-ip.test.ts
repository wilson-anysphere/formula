import assert from "node:assert/strict";
import http from "node:http";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";
import type { SyncServerConfig } from "../src/config.js";

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

test("auth:introspect caches results per (token, docId, clientIp)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-introspect-ip-cache-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const internalAdminToken = "internal-admin-token";
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
    if (provided !== internalAdminToken) {
      res.writeHead(403, { "content-type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: "forbidden" }));
      return;
    }

    const bodyText = await new Promise<string>((resolve) => {
      let data = "";
      req.setEncoding("utf8");
      req.on("data", (chunk) => {
        data += chunk;
      });
      req.on("end", () => resolve(data));
    });

    const body = JSON.parse(bodyText) as any;
    const token = body?.token;
    const docId = body?.docId;
    const clientIp = body?.clientIp;

    if (typeof token !== "string" || token.length === 0 || typeof docId !== "string" || docId.length === 0) {
      res.writeHead(400, { "content-type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: "invalid_request" }));
      return;
    }

    assert.ok(typeof clientIp === "string" && clientIp.length > 0);

    const key = `${token}\n${docId}\n${clientIp}`;
    hitsByKey.set(key, (hitsByKey.get(key) ?? 0) + 1);

    res.writeHead(200, { "content-type": "application/json" });
    res.end(JSON.stringify({ active: true, userId: "u1", orgId: "o1", role: "editor" }));
  });

  await new Promise<void>((resolve) => {
    introspectionServer.listen(0, "127.0.0.1", () => resolve());
  });
  t.after(async () => {
    await new Promise<void>((resolve, reject) => {
      introspectionServer.close((err) => (err ? reject(err) : resolve()));
    });
  });

  const addr = introspectionServer.address();
  assert.ok(addr && typeof addr !== "string");
  const introspectUrl = `http://127.0.0.1:${addr.port}`;

  const config: SyncServerConfig = {
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
      mode: "introspect",
      url: introspectUrl,
      token: internalAdminToken,
      cacheMs: 30_000,
      failOpen: false,
    },
    enforceRangeRestrictions: false,
    introspection: null,
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

  const logger = createLogger("silent");
  const server = createSyncServer(config, logger);
  await server.start();
  t.after(async () => {
    await server.stop();
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const token = "editor-token";
  const ipA = "1.2.3.4";
  const ipB = "5.6.7.8";

  const connectOnce = async (ip: string) => {
    const ws = new WebSocket(`${server.getWsUrl()}/${docName}?token=${encodeURIComponent(token)}`, {
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

  assert.equal(hitsByKey.get(`${token}\n${docName}\n${ipA}`), 1);
  assert.equal(hitsByKey.get(`${token}\n${docName}\n${ipB}`), 1);
});
