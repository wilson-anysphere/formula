import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import { startSyncServer, waitForCondition, waitForProviderSync } from "./test-helpers.ts";

async function waitForWsCloseWithTimeout(
  ws: WebSocket,
  timeoutMs: number
): Promise<{ code: number; reason: string }> {
  return await new Promise<{ code: number; reason: string }>((resolve, reject) => {
    let timer: NodeJS.Timeout | null = null;
    const onClose = (code: number, reason: Buffer) => {
      if (timer) clearTimeout(timer);
      timer = null;
      resolve({ code, reason: reason.toString("utf8") });
    };

    timer = setTimeout(() => {
      ws.off("close", onClose);
      reject(new Error("Timed out waiting for websocket close"));
    }, timeoutMs);
    timer.unref();

    ws.once("close", onClose);
  });
}

test("reserved history quotas: rejects branching commits beyond per-doc limit", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-branching-quota-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC: "2",
      SYNC_SERVER_MAX_VERSIONS_PER_DOC: "0",
    },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `quota-branching-${Math.random().toString(16).slice(2)}`;

  const docWriter = new Y.Doc();
  const docObserver = new Y.Doc();

  const providerWriter = new WebsocketProvider(server.wsUrl, docName, docWriter, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const providerObserver = new WebsocketProvider(server.wsUrl, docName, docObserver, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    providerWriter.destroy();
    providerObserver.destroy();
    docWriter.destroy();
    docObserver.destroy();
  });

  await waitForProviderSync(providerWriter);
  await waitForProviderSync(providerObserver);

  const commitsWriter = docWriter.getMap("branching:commits");

  docWriter.transact(() => {
    commitsWriter.set("c1", "one");
  });
  docWriter.transact(() => {
    commitsWriter.set("c2", "two");
  });

  await waitForCondition(() => docObserver.getMap("branching:commits").size === 2, 10_000);

  const writerWs = providerWriter.ws as WebSocket | undefined;
  assert.ok(writerWs, "Expected writer provider to have an underlying ws");
  const closePromise = waitForWsCloseWithTimeout(writerWs, 10_000);

  // This update should be rejected by the server quota guard.
  docWriter.transact(() => {
    commitsWriter.set("c3", "three");
  });

  const closed = await closePromise;
  assert.equal(closed.code, 1008);
  assert.match(closed.reason, /quota|history|reserved/i);
  providerWriter.destroy();

  // Ensure other collaborators never observe the rejected commit.
  await waitForCondition(() => docObserver.getMap("branching:commits").size === 2, 10_000);
  assert.equal(docObserver.getMap("branching:commits").has("c3"), false);

  // Ensure a fresh client also doesn't observe the rejected commit.
  const docFresh = new Y.Doc();
  const providerFresh = new WebsocketProvider(server.wsUrl, docName, docFresh, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    providerFresh.destroy();
    docFresh.destroy();
  });

  await waitForProviderSync(providerFresh);
  assert.equal(docFresh.getMap("branching:commits").size, 2);
  assert.equal(docFresh.getMap("branching:commits").has("c3"), false);

  // Metric is incremented (best-effort; accept >= 1 since clients may retry).
  await waitForCondition(
    async () => {
      const res = await fetch(`${server.httpUrl}/metrics`);
      if (!res.ok) return false;
      const body = await res.text();
      const match = body.match(
        /^sync_server_ws_reserved_root_quota_violations_total\{[^}]*kind="branching_commits"[^}]*\}\s+([0-9.]+)$/m
      );
      if (!match) return false;
      return Number(match[1]) >= 1;
    },
    5_000,
    100
  );
});

test("reserved history quotas: rejects versions beyond per-doc limit", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-versions-quota-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      SYNC_SERVER_MAX_BRANCHING_COMMITS_PER_DOC: "0",
      SYNC_SERVER_MAX_VERSIONS_PER_DOC: "2",
    },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `quota-versions-${Math.random().toString(16).slice(2)}`;

  const docWriter = new Y.Doc();
  const docObserver = new Y.Doc();

  const providerWriter = new WebsocketProvider(server.wsUrl, docName, docWriter, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const providerObserver = new WebsocketProvider(server.wsUrl, docName, docObserver, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    providerWriter.destroy();
    providerObserver.destroy();
    docWriter.destroy();
    docObserver.destroy();
  });

  await waitForProviderSync(providerWriter);
  await waitForProviderSync(providerObserver);

  const versionsWriter = docWriter.getMap("versions");

  docWriter.transact(() => {
    versionsWriter.set("v1", "one");
  });
  docWriter.transact(() => {
    versionsWriter.set("v2", "two");
  });

  await waitForCondition(() => docObserver.getMap("versions").size === 2, 10_000);

  const writerWs = providerWriter.ws as WebSocket | undefined;
  assert.ok(writerWs, "Expected writer provider to have an underlying ws");
  const closePromise = waitForWsCloseWithTimeout(writerWs, 10_000);

  // This update should be rejected by the server quota guard.
  docWriter.transact(() => {
    versionsWriter.set("v3", "three");
  });

  const closed = await closePromise;
  assert.equal(closed.code, 1008);
  assert.match(closed.reason, /quota|history|reserved/i);
  providerWriter.destroy();

  // Ensure other collaborators never observe the rejected version.
  await waitForCondition(() => docObserver.getMap("versions").size === 2, 10_000);
  assert.equal(docObserver.getMap("versions").has("v3"), false);

  // Ensure a fresh client also doesn't observe the rejected version.
  const docFresh = new Y.Doc();
  const providerFresh = new WebsocketProvider(server.wsUrl, docName, docFresh, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  t.after(() => {
    providerFresh.destroy();
    docFresh.destroy();
  });

  await waitForProviderSync(providerFresh);
  assert.equal(docFresh.getMap("versions").size, 2);
  assert.equal(docFresh.getMap("versions").has("v3"), false);

  await waitForCondition(
    async () => {
      const res = await fetch(`${server.httpUrl}/metrics`);
      if (!res.ok) return false;
      const body = await res.text();
      const match = body.match(
        /^sync_server_ws_reserved_root_quota_violations_total\{[^}]*kind="versions"[^}]*\}\s+([0-9.]+)$/m
      );
      if (!match) return false;
      return Number(match[1]) >= 1;
    },
    5_000,
    100
  );
});
