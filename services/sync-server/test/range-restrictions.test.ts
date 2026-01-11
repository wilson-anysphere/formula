import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import jwt from "jsonwebtoken";
import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";
import {
  FILE_HEADER_BYTES,
  hasFileHeader,
  scanLegacyRecords,
} from "../../../packages/collab/persistence/src/file-format.js";

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

async function waitForWsCloseWithTimeout(
  ws: WebSocket,
  timeoutMs: number
): Promise<{ code: number; reason: string }> {
  return await new Promise<{ code: number; reason: string }>((resolve, reject) => {
    let timer: NodeJS.Timeout | null = null;
    const onClose = (code: number, reason: Buffer) => {
      if (timer) clearTimeout(timer);
      timer = null;
      const reasonStr =
        typeof reason === "string"
          ? reason
          : Buffer.isBuffer(reason)
            ? reason.toString("utf8")
            : String(reason);
      resolve({ code, reason: reasonStr });
    };

    timer = setTimeout(() => {
      ws.off("close", onClose);
      reject(new Error("Timed out waiting for websocket close"));
    }, timeoutMs);

    ws.once("close", onClose);
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

function getCellsMap(doc: Y.Doc): Y.Map<unknown> {
  return doc.getMap<unknown>("cells");
}

function restrictionForB1(allowedEditorUserId: string) {
  return [
    {
      sheetId: "Sheet1",
      startRow: 0,
      startCol: 1,
      endRow: 0,
      endCol: 1,
      editAllowlist: [allowedEditorUserId],
    },
  ];
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

test("forbidden cell write is rejected and connection closed", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

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
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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

  assert.ok(providerWriter.ws, "Expected writer provider to have an underlying ws");
  const writerClose = waitForWsClose(providerWriter.ws);

  providerWriter.destroy();
  await writerClose;

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
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
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

test("rangeRestrictions: rejects invalid numeric cell keys (Infinity) without crashing", async (t) => {
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

  const seedToken = signJwt({ sub: "seed", docId: docName, role: "editor" });
  const attackerToken = signJwt({
    sub: "attacker",
    docId: docName,
    role: "editor",
    // Any non-empty array enables enforcement for this token.
    rangeRestrictions: [
      {
        range: { sheetId: "Sheet1", startRow: 0, endRow: 0, startCol: 0, endCol: 0 },
        editAllowlist: ["attacker"],
      },
    ],
  });

  const seedDoc = new Y.Doc();
  const attackerDoc = new Y.Doc();

  const seedProvider = new WebsocketProvider(server.wsUrl, docName, seedDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: seedToken },
  });
  const attackerProvider = new WebsocketProvider(server.wsUrl, docName, attackerDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: attackerToken },
  });

  t.after(() => {
    seedProvider.destroy();
    attackerProvider.destroy();
    seedDoc.destroy();
    attackerDoc.destroy();
  });

  await waitForProviderSync(seedProvider);
  await waitForProviderSync(attackerProvider);

  // Ensure we have something persisted to disk to validate later.
  setCellValue(seedDoc, "Sheet1:0:0", "seed");
  await waitForCondition(() => getCellValue(attackerDoc, "Sheet1:0:0") === "seed", 10_000);

  const hugeKey = `Sheet1:${"9".repeat(400)},1`;
  const ws = attackerProvider.ws;
  assert.ok(ws, "expected attacker provider to have an underlying ws");
  const closePromise = waitForWsCloseWithTimeout(ws, 10_000);

  // This key used to parse into `{ row: Infinity }` and could crash the server when the
  // permission check tried to validate row/col types. The server should now fail closed.
  setCellValue(attackerDoc, hugeKey, "evil");

  const closed = await closePromise;
  assert.equal(closed.code, 1008);

  await expectConditionToStayFalse(() => getCellsMap(seedDoc).has(hugeKey), 1_000);
  assert.equal(getCellsMap(seedDoc).has(hugeKey), false);

  const healthRes = await fetch(`${server.httpUrl}/healthz`);
  assert.equal(healthRes.status, 200);

  await waitForCondition(async () => {
    try {
      const persisted = await loadPersistedDoc(dataDir, docName);
      const ok =
        getCellValue(persisted, "Sheet1:0:0") === "seed" && getCellsMap(persisted).has(hugeKey) === false;
      persisted.destroy();
      return ok;
    } catch {
      return false;
    }
  }, 10_000);

  const persisted = await loadPersistedDoc(dataDir, docName);
  t.after(() => persisted.destroy());
  assert.equal(getCellValue(persisted, "Sheet1:0:0"), "seed");
  assert.equal(getCellsMap(persisted).has(hugeKey), false);
});

test("rangeRestrictions: allows edits outside protected ranges", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: { SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS: "1" },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;

  const ownerToken = signJwt({ sub: "owner", docId: docName, role: "editor" });
  const restrictedToken = signJwt({
    sub: "restricted",
    docId: docName,
    role: "editor",
    rangeRestrictions: restrictionForB1("owner"),
  });

  const seedDoc = new Y.Doc();
  const seedProvider = new WebsocketProvider(server.wsUrl, docName, seedDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: ownerToken },
  });

  const restrictedDoc = new Y.Doc();
  const restrictedProvider = new WebsocketProvider(server.wsUrl, docName, restrictedDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: restrictedToken },
  });

  let cleanedUp = false;
  const cleanup = () => {
    if (cleanedUp) return;
    cleanedUp = true;
    seedProvider.destroy();
    restrictedProvider.destroy();
    seedDoc.destroy();
    restrictedDoc.destroy();
  };
  t.after(cleanup);

  await waitForProviderSync(seedProvider);

  // Seed initial cells so subsequent writes are nested map updates (not new keys).
  setCellValue(seedDoc, "Sheet1:0:0", "initA");
  setCellValue(seedDoc, "Sheet1:0:1", "initB");

  await waitForProviderSync(restrictedProvider);
  await waitForCondition(() => getCellValue(restrictedDoc, "Sheet1:0:1") === "initB", 10_000);

  // A1 is outside the protected range (B1), so it should be accepted.
  setCellValue(restrictedDoc, "Sheet1:0:0", "allowedA");
  await waitForCondition(() => getCellValue(seedDoc, "Sheet1:0:0") === "allowedA", 10_000);

  cleanup();

  await waitForCondition(async () => {
    try {
      const persisted = await loadPersistedDoc(dataDir, docName);
      const ok = getCellValue(persisted, "Sheet1:0:0") === "allowedA";
      persisted.destroy();
      return ok;
    } catch {
      return false;
    }
  }, 10_000);

  const persisted = await loadPersistedDoc(dataDir, docName);
  t.after(() => persisted.destroy());

  assert.equal(getCellValue(persisted, "Sheet1:0:0"), "allowedA");
  assert.equal(getCellValue(persisted, "Sheet1:0:1"), "initB");
});

test("rangeRestrictions: blocks edits to protected cells and does not persist them", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: { SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS: "1" },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;

  const ownerToken = signJwt({ sub: "owner", docId: docName, role: "editor" });
  const restrictedToken = signJwt({
    sub: "restricted",
    docId: docName,
    role: "editor",
    rangeRestrictions: restrictionForB1("owner"),
  });

  const seedDoc = new Y.Doc();
  const seedProvider = new WebsocketProvider(server.wsUrl, docName, seedDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: ownerToken },
  });

  const restrictedDoc = new Y.Doc();
  const restrictedProvider = new WebsocketProvider(server.wsUrl, docName, restrictedDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: restrictedToken },
  });

  let cleanedUp = false;
  const cleanup = () => {
    if (cleanedUp) return;
    cleanedUp = true;
    seedProvider.destroy();
    restrictedProvider.destroy();
    seedDoc.destroy();
    restrictedDoc.destroy();
  };
  t.after(cleanup);

  await waitForProviderSync(seedProvider);
  setCellValue(seedDoc, "Sheet1:0:1", "initB");

  await waitForProviderSync(restrictedProvider);
  await waitForCondition(() => getCellValue(restrictedDoc, "Sheet1:0:1") === "initB", 10_000);

  const ws = (restrictedProvider as any).ws as WebSocket | undefined;
  assert.ok(ws, "expected restricted provider to have a ws");
  const closePromise = waitForWsCloseWithTimeout(ws, 10_000);

  // Edit B1 (existing cell); should be rejected and close the websocket.
  setCellValue(restrictedDoc, "Sheet1:0:1", "evilB");

  assert.equal((await closePromise).code, 1008);

  // Give the server a moment to (not) broadcast the forbidden update.
  await new Promise((r) => setTimeout(r, 250));
  assert.equal(getCellValue(seedDoc, "Sheet1:0:1"), "initB");

  cleanup();

  await waitForCondition(async () => {
    try {
      const persisted = await loadPersistedDoc(dataDir, docName);
      const ok = getCellValue(persisted, "Sheet1:0:1") === "initB";
      persisted.destroy();
      return ok;
    } catch {
      return false;
    }
  }, 10_000);

  const persisted = await loadPersistedDoc(dataDir, docName);
  t.after(() => persisted.destroy());

  assert.equal(getCellValue(persisted, "Sheet1:0:1"), "initB");
});

test("rangeRestrictions: legacy cell key formats cannot bypass protected ranges", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-"));

  let server = await startSyncServer({
    dataDir,
    auth: { mode: "jwt", secret: JWT_SECRET, audience: JWT_AUDIENCE },
    env: { SYNC_SERVER_ENFORCE_RANGE_RESTRICTIONS: "1" },
  });
  t.after(async () => {
    await server.stop();
  });
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const docName = `doc-${Math.random().toString(16).slice(2)}`;

  const ownerToken = signJwt({ sub: "owner", docId: docName, role: "editor" });
  const restrictedToken = signJwt({
    sub: "restricted",
    docId: docName,
    role: "editor",
    rangeRestrictions: restrictionForB1("owner"),
  });

  const seedDoc = new Y.Doc();
  const seedProvider = new WebsocketProvider(server.wsUrl, docName, seedDoc, {
    WebSocketPolyfill: WebSocket,
    disableBc: true,
    params: { token: ownerToken },
  });
  await waitForProviderSync(seedProvider);
  setCellValue(seedDoc, "Sheet1:0:1", "initB");

  let cleanedUp = false;
  const cleanup = () => {
    if (cleanedUp) return;
    cleanedUp = true;
    seedProvider.destroy();
    seedDoc.destroy();
  };
  t.after(cleanup);

  const attemptLegacyWrite = async (cellKey: string) => {
    const restrictedDoc = new Y.Doc();
    const restrictedProvider = new WebsocketProvider(server.wsUrl, docName, restrictedDoc, {
      WebSocketPolyfill: WebSocket,
      disableBc: true,
      params: { token: restrictedToken },
    });
    t.after(() => {
      restrictedProvider.destroy();
      restrictedDoc.destroy();
    });

    await waitForProviderSync(restrictedProvider);
    await waitForCondition(() => getCellValue(restrictedDoc, "Sheet1:0:1") === "initB", 10_000);

    const ws = (restrictedProvider as any).ws as WebSocket | undefined;
    assert.ok(ws, "expected restricted provider to have a ws");
    const closePromise = waitForWsCloseWithTimeout(ws, 10_000);

    setCellValue(restrictedDoc, cellKey, "evilLegacy");
    assert.equal((await closePromise).code, 1008);

    await new Promise((r) => setTimeout(r, 250));
    assert.equal(getCellValue(seedDoc, "Sheet1:0:1"), "initB");
  };

  // Attempt to write B1 using legacy key formats which must normalize to Sheet1:0:1
  // and still be rejected.
  await attemptLegacyWrite("r0c1");
  await attemptLegacyWrite("Sheet1:0,1");
  await attemptLegacyWrite(":0:1");
  await attemptLegacyWrite(":0,1");

  await waitForCondition(async () => {
    try {
      const persisted = await loadPersistedDoc(dataDir, docName);
      const ok =
        getCellValue(persisted, "Sheet1:0:1") === "initB" &&
        getCellsMap(persisted).has("r0c1") === false &&
        getCellsMap(persisted).has("Sheet1:0,1") === false &&
        getCellsMap(persisted).has(":0:1") === false &&
        getCellsMap(persisted).has(":0,1") === false;
      persisted.destroy();
      return ok;
    } catch {
      return false;
    }
  }, 10_000);

  const persisted = await loadPersistedDoc(dataDir, docName);
  t.after(() => persisted.destroy());

  assert.equal(getCellValue(persisted, "Sheet1:0:1"), "initB");
  assert.equal(getCellsMap(persisted).has("r0c1"), false);
  assert.equal(getCellsMap(persisted).has("Sheet1:0,1"), false);
  assert.equal(getCellsMap(persisted).has(":0:1"), false);
  assert.equal(getCellsMap(persisted).has(":0,1"), false);
});
