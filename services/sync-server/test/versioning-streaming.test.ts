import assert from "node:assert/strict";
import crypto from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";

import { CollabVersioning } from "../../../packages/collab/versioning/src/index.ts";
import { YjsVersionStore } from "../../../packages/versioning/src/store/yjsVersionStore.js";

import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";
import { startSyncServer, waitForCondition, waitForProviderSync } from "./test-helpers.ts";

test("streams large version snapshots without exceeding SYNC_SERVER_MAX_MESSAGE_BYTES", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-versioning-stream-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      // Small enough that a single snapshot write would reliably exceed the limit,
      // but large enough to avoid flakiness due to protocol overhead.
      SYNC_SERVER_MAX_MESSAGE_BYTES: String(128 * 1024),
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "versioning-streaming-doc";
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  const providerA = new WebsocketProvider(server.wsUrl, docName, docA, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const providerB = new WebsocketProvider(server.wsUrl, docName, docB, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    providerA.destroy();
    providerB.destroy();
    docA.destroy();
    docB.destroy();
  });

  await waitForProviderSync(providerA);
  await waitForProviderSync(providerB);

  let saw1009 = false;
  const attachCloseListener = (provider: any) => {
    const ws = provider?.ws;
    if (!ws || typeof ws.on !== "function") return;
    ws.on("close", (code: number) => {
      if (code === 1009) saw1009 = true;
    });
  };
  attachCloseListener(providerA as any);
  attachCloseListener(providerB as any);

  // Build a large doc state in many small updates so we don't trip the server
  // limit during normal editing.
  const bigText = docA.getText("bigText");
  const chunk = "x".repeat(16 * 1024);
  const repeats = 24; // ~384KB of text content.
  for (let i = 0; i < repeats; i += 1) {
    docA.transact(() => {
      bigText.insert(bigText.length, chunk);
    }, "test-bigtext");
  }

  await waitForCondition(() => docB.getText("bigText").length === bigText.length, 10_000, 50);

  const storeA = new YjsVersionStore({
    doc: docA,
    writeMode: "stream",
    // Split the snapshot bytes so each Yjs update stays well under 128KB.
    chunkSize: 32 * 1024,
    maxChunksPerTransaction: 2,
  });

  const versioningA = new CollabVersioning({
    session: { doc: docA } as any,
    store: storeA,
    autoStart: false,
  });

  const created = await versioningA.createSnapshot({ description: "big snapshot" });
  assert.ok(
    created.snapshot.byteLength > 128 * 1024,
    `expected snapshot bytes > 128KB, got ${created.snapshot.byteLength}`
  );

  // Wait for the streamed record to fully arrive on the other client.
  await waitForCondition(() => {
    const raw = docB.getMap("versions").get(created.id) as any;
    return Boolean(raw?.get?.("snapshotComplete") === true);
  }, 20_000, 50);

  const storeB = new YjsVersionStore({ doc: docB });
  const listed = await storeB.listVersions();
  assert.ok(listed.some((v) => v.id === created.id));

  const fetched = await storeB.getVersion(created.id);
  assert.ok(fetched);
  assert.equal(fetched.id, created.id);

  // Compare via hashes to keep the assertion cheap for large payloads.
  const hash = (bytes: Uint8Array) =>
    crypto.createHash("sha256").update(Buffer.from(bytes)).digest("hex");
  assert.equal(hash(fetched.snapshot), hash(created.snapshot));

  const metricsRes = await fetch(`${server.httpUrl}/metrics`);
  assert.equal(metricsRes.ok, true);
  const metricsText = await metricsRes.text();
  const match = metricsText.match(
    /^sync_server_ws_messages_too_large_total(?:\{[^}]*\})?\s+([0-9.]+)$/m
  );
  assert.ok(match, "expected sync_server_ws_messages_too_large_total metric");
  assert.equal(Number(match[1]), 0);
  assert.equal(saw1009, false);
});

test("CollabVersioning defaults stream large snapshots under typical SYNC_SERVER_MAX_MESSAGE_BYTES", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-versioning-default-stream-"));

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      // Mimic a more typical deployment limit. The default CollabVersioning store
      // should stream snapshot chunks so each update stays under this threshold.
      SYNC_SERVER_MAX_MESSAGE_BYTES: String(1024 * 1024),
    },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = "versioning-default-streaming-doc";
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  const providerA = new WebsocketProvider(server.wsUrl, docName, docA, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });
  const providerB = new WebsocketProvider(server.wsUrl, docName, docB, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: "test-token" },
  });

  t.after(() => {
    providerA.destroy();
    providerB.destroy();
    docA.destroy();
    docB.destroy();
  });

  await waitForProviderSync(providerA);
  await waitForProviderSync(providerB);

  let saw1009 = false;
  const attachCloseListener = (provider: any) => {
    const ws = provider?.ws;
    if (!ws || typeof ws.on !== "function") return;
    ws.on("close", (code: number) => {
      if (code === 1009) saw1009 = true;
    });
  };
  attachCloseListener(providerA as any);
  attachCloseListener(providerB as any);

  // Build a large doc state in many small updates so we don't trip the server
  // limit during normal editing.
  const bigText = docA.getText("bigText");
  const chunk = "x".repeat(16 * 1024);
  const repeats = 96; // ~1.5MB of text content.
  for (let i = 0; i < repeats; i += 1) {
    docA.transact(() => {
      bigText.insert(bigText.length, chunk);
    }, "test-bigtext");
  }

  await waitForCondition(() => docB.getText("bigText").length === bigText.length, 30_000, 50);

  // No explicit store config: this exercises CollabVersioning's default store settings.
  const versioningA = new CollabVersioning({
    session: { doc: docA } as any,
    autoStart: false,
  });
  t.after(() => versioningA.destroy());

  const created = await versioningA.createSnapshot({ description: "big snapshot (default store)" });
  assert.ok(
    created.snapshot.byteLength > 1024 * 1024,
    `expected snapshot bytes > 1MB, got ${created.snapshot.byteLength}`,
  );

  // Wait for the streamed record to fully arrive on the other client.
  await waitForCondition(() => {
    const raw = docB.getMap("versions").get(created.id) as any;
    return Boolean(raw?.get?.("snapshotComplete") === true);
  }, 30_000, 50);

  const storeB = new YjsVersionStore({ doc: docB });
  const fetched = await storeB.getVersion(created.id);
  assert.ok(fetched);

  // Compare via hashes to keep the assertion cheap for large payloads.
  const hash = (bytes: Uint8Array) =>
    crypto.createHash("sha256").update(Buffer.from(bytes)).digest("hex");
  assert.equal(hash(fetched.snapshot), hash(created.snapshot));

  const metricsRes = await fetch(`${server.httpUrl}/metrics`);
  assert.equal(metricsRes.ok, true);
  const metricsText = await metricsRes.text();
  const match = metricsText.match(
    /^sync_server_ws_messages_too_large_total(?:\{[^}]*\})?\s+([0-9.]+)$/m,
  );
  assert.ok(match, "expected sync_server_ws_messages_too_large_total metric");
  assert.equal(Number(match[1]), 0);
  assert.equal(saw1009, false);
});
