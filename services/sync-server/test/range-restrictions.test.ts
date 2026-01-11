import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";
import * as Y from "yjs";
import { WebsocketProvider } from "y-websocket";

import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

const JWT_SECRET = "test-secret";
const JWT_AUDIENCE = "formula-sync";

function signJwt(payload: Record<string, unknown>): string {
  return jwt.sign(payload, JWT_SECRET, {
    algorithm: "HS256",
    audience: JWT_AUDIENCE,
  });
}

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

function getCellValue(doc: Y.Doc, cellKey: string): unknown {
  const cell = doc.getMap("cells").get(cellKey) as any;
  if (!cell || typeof cell !== "object") return null;
  if (typeof cell.get !== "function") return null;
  return cell.get("value") ?? null;
}

function getCellModifiedBy(doc: Y.Doc, cellKey: string): unknown {
  const cell = doc.getMap("cells").get(cellKey) as any;
  if (!cell || typeof cell !== "object") return null;
  if (typeof cell.get !== "function") return null;
  return cell.get("modifiedBy") ?? null;
}

test("forbidden cell write is rejected and connection closed", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: { SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS: "1" },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const cellKey = "Sheet1:0:0";

  const editorToken = signJwt({
    sub: "editor",
    docId: docName,
    orgId: "o1",
    role: "editor",
  });

  const restrictedToken = signJwt({
    sub: "restricted",
    docId: docName,
    orgId: "o1",
    role: "editor",
    rangeRestrictions: [
      {
        range: { sheetId: "Sheet1", startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        // Deny edits by excluding this user from the allowlist.
        editAllowlist: ["someone-else"],
      },
    ],
  });

  const docEditor = new Y.Doc();
  const docRestricted = new Y.Doc();

  const providerEditor = new WebsocketProvider(server.wsUrl, docName, docEditor, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });
  const providerRestricted = new WebsocketProvider(server.wsUrl, docName, docRestricted, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: restrictedToken },
  });

  t.after(() => {
    providerEditor.destroy();
    providerRestricted.destroy();
    docEditor.destroy();
    docRestricted.destroy();
  });

  await waitForProviderSync(providerEditor);
  await waitForProviderSync(providerRestricted);

  assert.ok(providerRestricted.ws, "Expected restricted provider to have an underlying ws");
  const closePromise = waitForWsClose(providerRestricted.ws);

  docRestricted.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "evil");
    docRestricted.getMap("cells").set(cellKey, cell);
  });

  const closed = await closePromise;
  assert.equal(closed.code, 1008);
  assert.match(closed.reason, /permission|range restrictions|unparseable/i);

  await expectConditionToStayFalse(
    () => getCellValue(docEditor, cellKey) === "evil",
    1_000
  );
  assert.notEqual(getCellValue(docEditor, cellKey), "evil");

  // Ensure the forbidden update was not persisted.
  const docEditor2 = new Y.Doc();
  const providerEditor2 = new WebsocketProvider(server.wsUrl, docName, docEditor2, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: editorToken },
  });

  t.after(() => {
    providerEditor2.destroy();
    docEditor2.destroy();
  });

  await waitForProviderSync(providerEditor2);
  await waitForCondition(() => providerEditor2.synced === true, 10_000);
  assert.notEqual(getCellValue(docEditor2, cellKey), "evil");
});

test("allowed cell write is accepted and syncs", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: { SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS: "1" },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const cellKey = "Sheet1:0:0";

  const writerToken = signJwt({
    sub: "writer",
    docId: docName,
    orgId: "o1",
    role: "editor",
    rangeRestrictions: [
      {
        range: { sheetId: "Sheet1", startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        editAllowlist: ["writer"],
      },
    ],
  });
  const readerToken = signJwt({
    sub: "reader",
    docId: docName,
    orgId: "o1",
    role: "editor",
  });

  const docWriter = new Y.Doc();
  const docReader = new Y.Doc();

  const providerWriter = new WebsocketProvider(server.wsUrl, docName, docWriter, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: writerToken },
  });
  const providerReader = new WebsocketProvider(server.wsUrl, docName, docReader, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: readerToken },
  });

  t.after(() => {
    providerWriter.destroy();
    providerReader.destroy();
    docWriter.destroy();
    docReader.destroy();
  });

  await waitForProviderSync(providerWriter);
  await waitForProviderSync(providerReader);

  docWriter.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "ok");
    // Attempt to spoof audit metadata; server should rewrite to authenticated userId.
    cell.set("modifiedBy", "spoofed");
    docWriter.getMap("cells").set(cellKey, cell);
  });

  await waitForCondition(
    () =>
      getCellValue(docReader, cellKey) === "ok" &&
      getCellModifiedBy(docReader, cellKey) === "writer",
    10_000
  );
  assert.equal(getCellValue(docReader, cellKey), "ok");
  assert.equal(getCellModifiedBy(docReader, cellKey), "writer");
});

test("allowed offline edit syncs on reconnect (shadow state seeded from server doc)", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: { SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS: "1" },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;
  const cellKey = "Sheet1:0:0";

  const writerToken = signJwt({
    sub: "writer",
    docId: docName,
    orgId: "o1",
    role: "editor",
    rangeRestrictions: [
      {
        range: { sheetId: "Sheet1", startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        editAllowlist: ["writer"],
      },
    ],
  });
  const observerToken = signJwt({
    sub: "observer",
    docId: docName,
    orgId: "o1",
    role: "editor",
  });

  const docWriter = new Y.Doc();
  const providerWriter = new WebsocketProvider(server.wsUrl, docName, docWriter, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: writerToken },
  });

  t.after(() => {
    providerWriter.destroy();
    docWriter.destroy();
  });

  await waitForProviderSync(providerWriter);

  docWriter.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "base");
    docWriter.getMap("cells").set(cellKey, cell);
  });

  providerWriter.destroy();
  await new Promise((r) => setTimeout(r, 250));

  // Offline edit: modify an existing cell *before* reconnecting.
  docWriter.transact(() => {
    const cell = docWriter.getMap("cells").get(cellKey) as any;
    assert.ok(cell, "expected cell to exist in offline doc");
    cell.set("value", "offline");
    cell.set("modifiedBy", "spoofed");
  });

  const providerReconnect = new WebsocketProvider(server.wsUrl, docName, docWriter, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: writerToken },
  });

  const docObserver = new Y.Doc();
  const providerObserver = new WebsocketProvider(server.wsUrl, docName, docObserver, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: observerToken },
  });

  t.after(() => {
    providerReconnect.destroy();
    providerObserver.destroy();
    docObserver.destroy();
  });

  await waitForProviderSync(providerReconnect);
  await waitForProviderSync(providerObserver);

  await waitForCondition(
    () =>
      getCellValue(docObserver, cellKey) === "offline" &&
      getCellModifiedBy(docObserver, cellKey) === "writer",
    10_000
  );

  assert.equal(getCellValue(docObserver, cellKey), "offline");
  assert.equal(getCellModifiedBy(docObserver, cellKey), "writer");
});

test("strict mode rejects updates when cell keys cannot be parsed", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: { SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS: "1" },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;

  const attackerToken = signJwt({
    sub: "attacker",
    docId: docName,
    orgId: "o1",
    role: "editor",
    // Any non-empty array enables enforcement for this token.
    rangeRestrictions: [
      {
        range: { sheetId: "Sheet1", startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        editAllowlist: ["attacker"],
      },
    ],
  });

  const observerToken = signJwt({
    sub: "observer",
    docId: docName,
    orgId: "o1",
    role: "editor",
  });

  const docAttacker = new Y.Doc();
  const docObserver = new Y.Doc();

  const attackerProvider = new WebsocketProvider(server.wsUrl, docName, docAttacker, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: attackerToken },
  });
  const observerProvider = new WebsocketProvider(server.wsUrl, docName, docObserver, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: observerToken },
  });

  t.after(() => {
    attackerProvider.destroy();
    observerProvider.destroy();
    docAttacker.destroy();
    docObserver.destroy();
  });

  await waitForProviderSync(attackerProvider);
  await waitForProviderSync(observerProvider);

  assert.ok(attackerProvider.ws, "Expected attacker provider to have an underlying ws");
  const closePromise = waitForWsClose(attackerProvider.ws);

  docAttacker.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "evil");
    docAttacker.getMap("cells").set("not-a-cell-key", cell);
  });

  const closed = await closePromise;
  assert.equal(closed.code, 1008);
  assert.match(closed.reason, /unparseable|range restrictions|permission/i);

  await expectConditionToStayFalse(
    () => docObserver.getMap("cells").has("not-a-cell-key"),
    1_000
  );
  assert.equal(docObserver.getMap("cells").has("not-a-cell-key"), false);
});
