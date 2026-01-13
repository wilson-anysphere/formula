import assert from "node:assert/strict";
import crypto from "node:crypto";
import net from "node:net";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";

import type { SyncServerConfig } from "../src/config.js";
import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";

function createConfig(dataDir: string, secret: string): SyncServerConfig {
  return {
    host: "127.0.0.1",
    port: 0,
    trustProxy: false,
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
      secret,
      audience: "formula-sync",
      requireSub: false,
      requireExp: false,
    },
    enforceRangeRestrictions: false,
    introspection: null,
    internalAdminToken: null,
    retention: { ttlMs: 0, sweepIntervalMs: 0, tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000 },
    limits: {
      maxConnections: 100,
      maxConnectionsPerIp: 25,
      maxConnectionsPerDoc: 0,
      maxConnAttemptsPerWindow: 100,
      connAttemptWindowMs: 60_000,
      maxMessageBytes: 2 * 1024 * 1024,
      maxMessagesPerWindow: 2_000,
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
}

async function rawWebSocketUpgradeStatus(opts: {
  host: string;
  port: number;
  pathWithQuery: string;
  method?: string;
}): Promise<number> {
  return await new Promise((resolve, reject) => {
    const socket = net.connect(opts.port, opts.host);
    socket.setTimeout(5_000);

    socket.on("timeout", () => {
      socket.destroy();
      reject(new Error("Timed out waiting for upgrade response"));
    });
    socket.on("error", reject);

    socket.on("connect", () => {
      const key = crypto.randomBytes(16).toString("base64");
      const method = opts.method ?? "GET";
      const request = [
        `${method} ${opts.pathWithQuery} HTTP/1.1`,
        `Host: ${opts.host}:${opts.port}`,
        "Upgrade: websocket",
        "Connection: Upgrade",
        `Sec-WebSocket-Key: ${key}`,
        "Sec-WebSocket-Version: 13",
        "",
        "",
      ].join("\r\n");
      socket.write(request);
    });

    let buffer = "";
    socket.setEncoding("utf8");
    socket.on("data", (chunk) => {
      buffer += chunk;
      const headerEnd = buffer.indexOf("\r\n\r\n");
      if (headerEnd < 0) return;

      const header = buffer.slice(0, headerEnd);
      const statusLine = header.split("\r\n")[0] ?? "";
      const match = statusLine.match(/^HTTP\/1\.1\s+(\d+)/i);
      if (!match) {
        socket.destroy();
        reject(new Error(`Unexpected response: ${statusLine}`));
        return;
      }

      const statusCode = Number(match[1]);
      socket.destroy();
      resolve(statusCode);
    });
  });
}

test("docName extraction does not normalize dot segments during websocket upgrade", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-docname-dot-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const secret = "test-secret";
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(dataDir, secret), logger);
  const { port } = await server.start();
  t.after(async () => {
    await server.stop();
  });

  // If docName parsing normalized `/a/../b` to `b`, this token would be accepted.
  // We expect the server to treat the doc name as the literal `a/../b` (matching
  // y-websocket) and reject this token.
  const token = jwt.sign(
    { sub: "u1", docId: "b", orgId: "o1", role: "editor" },
    secret,
    { algorithm: "HS256", audience: "formula-sync" }
  );

  const status = await rawWebSocketUpgradeStatus({
    host: "127.0.0.1",
    port,
    pathWithQuery: `/a/../b?token=${encodeURIComponent(token)}`,
  });
  assert.equal(status, 403);
});

test("websocket upgrade only accepts GET requests", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-docname-method-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const secret = "test-secret";
  const logger = createLogger("silent");
  const server = createSyncServer(createConfig(dataDir, secret), logger);
  const { port } = await server.start();
  t.after(async () => {
    await server.stop();
  });

  const docName = "method-doc";
  const token = jwt.sign(
    { sub: "u1", docId: docName, orgId: "o1", role: "editor" },
    secret,
    { algorithm: "HS256", audience: "formula-sync" }
  );

  const status = await rawWebSocketUpgradeStatus({
    host: "127.0.0.1",
    port,
    method: "POST",
    pathWithQuery: `/${docName}?token=${encodeURIComponent(token)}`,
  });
  assert.equal(status, 405);
});
