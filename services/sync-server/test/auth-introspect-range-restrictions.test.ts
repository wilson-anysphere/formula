import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import http from "node:http";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import {
  FILE_HEADER_BYTES,
  hasFileHeader,
  scanLegacyRecords,
} from "../../../packages/collab/persistence/src/file-format.js";
import { createLogger } from "../src/logger.js";
import { createSyncServer } from "../src/server.js";
import type { SyncServerConfig } from "../src/config.js";
import { waitForCondition, waitForProviderSync } from "./test-helpers.ts";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

async function waitForWsClose(ws: WebSocket): Promise<{ code: number; reason: string }> {
  return await new Promise<{ code: number; reason: string }>((resolve) => {
    ws.once("close", (code, reason) => {
      const reasonStr =
        typeof reason === "string"
          ? reason
          : Buffer.isBuffer(reason)
            ? reason.toString("utf8")
            : String(reason);
      resolve({ code, reason: reasonStr });
    });
  });
}

function getCellValue(doc: Y.Doc, cellKey: string): unknown {
  const cell = doc.getMap("cells").get(cellKey) as any;
  if (!cell || typeof cell !== "object") return null;
  if (typeof cell.get !== "function") return null;
  return cell.get("value") ?? null;
}

function setCellValue(doc: Y.Doc, cellKey: string, value: unknown): void {
  doc.transact(() => {
    const cells = doc.getMap<unknown>("cells");
    let cell = cells.get(cellKey);
    if (!(cell instanceof Y.Map)) {
      cell = new Y.Map();
      cells.set(cellKey, cell);
    }
    (cell as Y.Map<unknown>).set("value", value);
  });
}

function persistedDocPath(dataDir: string, docName: string): string {
  const docHash = createHash("sha256").update(docName).digest("hex");
  return path.join(dataDir, `${docHash}.yjs`);
}

async function loadPersistedDoc(dataDir: string, docName: string): Promise<Y.Doc> {
  const filePath = persistedDocPath(dataDir, docName);
  const data = await readFile(filePath);
  const doc = new Y.Doc();
  const updates = hasFileHeader(data)
    ? scanLegacyRecords(data, FILE_HEADER_BYTES).updates
    : scanLegacyRecords(data).updates;
  for (const update of updates) {
    Y.applyUpdate(doc, update);
  }
  return doc;
}

test("auth:introspect enforces rangeRestrictions and rejects forbidden cell writes", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-introspect-range-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const internalAdminToken = "internal-admin-token";

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
      res.end(JSON.stringify({ active: false, reason: "forbidden" }));
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

    let body: any;
    try {
      body = JSON.parse(bodyText);
    } catch {
      res.writeHead(400, { "content-type": "application/json" });
      res.end(JSON.stringify({ active: false, reason: "invalid_request" }));
      return;
    }

    const token = body?.token;
    const docId = body?.docId;
    if (typeof token !== "string" || token.length === 0 || typeof docId !== "string" || docId.length === 0) {
      res.writeHead(400, { "content-type": "application/json" });
      res.end(JSON.stringify({ active: false, reason: "invalid_request" }));
      return;
    }

    if (token !== "owner-token" && token !== "restricted-token") {
      res.writeHead(403, { "content-type": "application/json" });
      res.end(JSON.stringify({ active: false, reason: "forbidden" }));
      return;
    }

    const userId = token === "owner-token" ? "owner" : "restricted";

    res.writeHead(200, { "content-type": "application/json" });
    res.end(
      JSON.stringify({
        active: true,
        userId,
        orgId: "o1",
        role: "editor",
        ...(token === "restricted-token"
          ? {
              // Restrict edits to A1 to only the "owner" user.
              rangeRestrictions: [
                {
                  sheetId: "Sheet1",
                  startRow: 0,
                  startCol: 0,
                  endRow: 0,
                  endCol: 0,
                  editAllowlist: ["owner"],
                },
              ],
            }
          : {}),
      })
    );
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
      maxQueueDepthPerDoc: 0,
      maxQueueDepthTotal: 0,
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
    enforceRangeRestrictions: true,
    introspection: null,
    internalAdminToken: null,
    retention: { ttlMs: 0, sweepIntervalMs: 0, tombstoneTtlMs: 7 * 24 * 60 * 60 * 1000 },
    limits: {
      maxUrlBytes: 8192,
      maxTokenBytes: 4096,
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

  let stopped = false;
  t.after(async () => {
    if (stopped) return;
    stopped = true;
    await server.stop();
  });

  await server.start();

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const cellKey = "Sheet1:0:0";

  const docOwner = new Y.Doc();
  const docRestricted = new Y.Doc();

  const providerOwner = new WebsocketProvider(server.getWsUrl(), docName, docOwner, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "owner-token" },
  });
  const providerRestricted = new WebsocketProvider(server.getWsUrl(), docName, docRestricted, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "restricted-token" },
  });

  t.after(() => {
    providerOwner.destroy();
    providerRestricted.destroy();
    docOwner.destroy();
    docRestricted.destroy();
  });

  await waitForProviderSync(providerOwner);
  await waitForProviderSync(providerRestricted);

  // Seed the doc so the restricted write targets an existing cell.
  setCellValue(docOwner, cellKey, "base");
  await waitForCondition(() => getCellValue(docRestricted, cellKey) === "base", 10_000);

  assert.ok(providerRestricted.ws, "Expected restricted provider to have an underlying ws");
  const closePromise = waitForWsClose(providerRestricted.ws);

  // Attempt a forbidden write by the restricted user; should close the socket and not persist.
  setCellValue(docRestricted, cellKey, "evil");

  const closed = await closePromise;
  assert.equal(closed.code, 1008);
  assert.match(closed.reason, /permission|range restrictions|unparseable/i);

  // The forbidden update should not be applied to the shared doc state.
  await waitForCondition(() => getCellValue(docOwner, cellKey) === "base", 10_000);
  assert.equal(getCellValue(docOwner, cellKey), "base");

  // Stop the server to flush persistence before reading the on-disk snapshot.
  await server.stop();
  stopped = true;

  const persisted = await loadPersistedDoc(dataDir, docName);
  t.after(() => persisted.destroy());
  assert.equal(getCellValue(persisted, cellKey), "base");
});
