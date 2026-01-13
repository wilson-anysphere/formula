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
import { waitForCondition, waitForProviderSync } from "./test-helpers.ts";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

async function expectConditionToStayFalse(
  condition: () => boolean,
  timeoutMs: number
): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    assert.equal(condition(), false);
    await new Promise((r) => setTimeout(r, 25));
  }
}

test("auth:introspect enforces roles and caches introspection results", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-introspect-"));

  const internalAdminToken = "internal-admin-token";
  const hitsByKey = new Map<string, number>();
  const calls: Array<{ token: string; docId: string; clientIp: unknown; userAgent: unknown }> = [];

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

    let body: any;
    try {
      body = JSON.parse(bodyText);
    } catch {
      res.writeHead(400, { "content-type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: "forbidden" }));
      return;
    }

    const token = body?.token;
    const docId = body?.docId;
    const clientIp = body?.clientIp;
    const userAgent = body?.userAgent;

    if (typeof token !== "string" || token.length === 0 || typeof docId !== "string" || docId.length === 0) {
      res.writeHead(403, { "content-type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: "forbidden" }));
      return;
    }

    const key = `${token}\n${docId}`;
    hitsByKey.set(key, (hitsByKey.get(key) ?? 0) + 1);
    calls.push({ token, docId, clientIp, userAgent });

    const role =
      token === "editor-token"
        ? "editor"
        : token === "viewer-token"
          ? "viewer"
          : null;
    const userId =
      token === "editor-token"
        ? "editor"
        : token === "viewer-token"
          ? "viewer"
          : null;

    if (!role || !userId) {
      res.writeHead(403, { "content-type": "application/json" });
      res.end(JSON.stringify({ ok: false, error: "forbidden" }));
      return;
    }

    res.writeHead(200, { "content-type": "application/json" });
    const successBody =
      token === "editor-token"
        ? { active: true, userId, orgId: "o1", role }
        : { ok: true, userId, orgId: "o1", role };
    res.end(JSON.stringify(successBody));
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
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const otherDocName = `doc-${Math.random().toString(16).slice(2)}`;
  const editorToken = "editor-token";
  const viewerToken = "viewer-token";

  const docEditor = new Y.Doc();
  const docViewer = new Y.Doc();

  const providerEditor = new WebsocketProvider(server.getWsUrl(), docName, docEditor, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  const providerViewer = new WebsocketProvider(server.getWsUrl(), docName, docViewer, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: viewerToken },
  });

  t.after(() => {
    providerEditor.destroy();
    providerViewer.destroy();
    docEditor.destroy();
    docViewer.destroy();
  });

  await waitForProviderSync(providerEditor);
  await waitForProviderSync(providerViewer);

  docEditor.getText("t").insert(0, "hello");

  await waitForCondition(() => docViewer.getText("t").toString() === "hello", 10_000);
  assert.equal(docViewer.getText("t").toString(), "hello");

  docViewer.getText("t").insert(docViewer.getText("t").length, "evil");
  await expectConditionToStayFalse(
    () => docEditor.getText("t").toString().includes("evil"),
    1_000
  );

  providerViewer.destroy();
  docViewer.destroy();

  const docEditor2 = new Y.Doc();
  const providerEditor2 = new WebsocketProvider(server.getWsUrl(), docName, docEditor2, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  t.after(() => {
    providerEditor2.destroy();
    docEditor2.destroy();
  });

  await waitForProviderSync(providerEditor2);
  await waitForCondition(() => docEditor2.getText("t").toString() === "hello", 10_000);

  const doc1EditorKey = `${editorToken}\n${docName}`;
  const doc1ViewerKey = `${viewerToken}\n${docName}`;

  assert.equal(hitsByKey.get(doc1EditorKey), 1);
  assert.equal(hitsByKey.get(doc1ViewerKey), 1);

  // The same token should be usable across multiple docs; caching is scoped per
  // (token, docId, clientIp).
  const docEditor3 = new Y.Doc();
  const providerEditor3 = new WebsocketProvider(server.getWsUrl(), otherDocName, docEditor3, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  t.after(() => {
    providerEditor3.destroy();
    docEditor3.destroy();
  });
  await waitForProviderSync(providerEditor3);

  assert.equal(hitsByKey.get(`${editorToken}\n${otherDocName}`), 1);

  for (const call of calls) {
    assert.ok(typeof call.clientIp === "string" && call.clientIp.length > 0);
    assert.notEqual(call.clientIp, "unknown");
  }

  const metricsRes = await fetch(`${server.getHttpUrl()}/metrics`);
  assert.equal(metricsRes.status, 200);
  const metricsText = await metricsRes.text();
  const match = metricsText.match(
    /sync_server_introspection_request_duration_ms\{(?=[^}]*path="auth_mode")(?=[^}]*result="ok")[^}]*\}\s+([0-9eE+\-.]+)/
  );
  assert.ok(
    match,
    `Expected sync_server_introspection_request_duration_ms{path="auth_mode",result="ok"} in /metrics:\n${metricsText}`
  );
  assert.ok(Number(match[1]) > 0);
});
