import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

import WebSocket from "ws";
import { WebsocketProvider, Y } from "./yjs-interop.ts";

import {
  startSyncServer,
  waitForCondition,
  waitForProviderSync,
} from "./test-helpers.ts";

import { diffDocumentStates } from "../../../packages/versioning/branches/src/patch.js";
import { emptyDocumentState, normalizeDocumentState } from "../../../packages/versioning/branches/src/state.js";
import { YjsBranchStore } from "../../../packages/versioning/branches/src/store/YjsBranchStore.js";

/**
 * Deterministic pseudo-random bytes (avoid crypto + avoid huge repeated strings
 * that gzip would compress too well).
 */
function pseudoRandomBase64(byteLength: number): string {
  const out = new Uint8Array(byteLength);
  let x = 0x12345678;
  for (let i = 0; i < out.length; i += 1) {
    x = (1103515245 * x + 12345) >>> 0;
    out[i] = x & 0xff;
  }
  return Buffer.from(out).toString("base64");
}

function makeInitialState() {
  const state = emptyDocumentState();
  state.sheets.order = ["Sheet1"];
  state.sheets.metaById = {
    Sheet1: { id: "Sheet1", name: "Sheet1", view: { frozenRows: 0, frozenCols: 0 } },
  };
  state.cells = { Sheet1: {} };
  return state;
}

test("branching: large commits can sync under maxMessageBytes via gzip-chunks payloads", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-branching-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const server = await startSyncServer({
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
    env: {
      // Small enough that a single JSON commit payload would reliably fail, but
      // still large enough for reasonable chunk sizes.
      SYNC_SERVER_MAX_MESSAGE_BYTES: "32768",
    },
  });
  t.after(async () => {
    await server.stop();
  });

  const docName = "branching-large-commit-doc";
  const docId = docName;

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

  const closeCodes: number[] = [];
  const recordClose = (code: unknown) => {
    if (typeof code === "number") closeCodes.push(code);
  };

  const attachProviderWsCloseListener = () => {
    const ws = (providerA as any).ws as WebSocket | undefined;
    if (!ws) return;
    const marker = "__branchingCloseListenerInstalled";
    if ((ws as any)[marker]) return;
    (ws as any)[marker] = true;
    ws.on("close", (code) => recordClose(code));
  };

  // y-websocket emits both `connection-close` and a raw `ws` close event; capture both.
  (providerA as any).on?.("connection-close", (event: any) => recordClose(event?.code));
  (providerA as any).on?.("status", (event: any) => {
    const status = typeof event === "string" ? event : event?.status;
    if (status === "connected") attachProviderWsCloseListener();
  });

  await waitForProviderSync(providerA);
  await waitForProviderSync(providerB);
  attachProviderWsCloseListener();

  const storeA = new YjsBranchStore({
    // yjs-interop uses the CJS build; YjsBranchStore supports mixed module instances.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ydoc: docA as any,
    payloadEncoding: "gzip-chunks",
    chunkSize: 4096,
    maxChunksPerTransaction: 1,
    snapshotEveryNCommits: 1,
  });
  const storeB = new YjsBranchStore({
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    ydoc: docB as any,
    payloadEncoding: "gzip-chunks",
    chunkSize: 4096,
    maxChunksPerTransaction: 1,
    snapshotEveryNCommits: 1,
  });

  const actor = { userId: "u1", role: "owner" as const };
  await storeA.ensureDocument(docId, actor, makeInitialState());

  await waitForCondition(async () => Boolean(await storeB.getBranch(docId, "main")), 10_000);

  const mainA = await storeA.getBranch(docId, "main");
  assert.ok(mainA);

  const current = await storeA.getDocumentStateAtCommit(mainA.headCommitId);
  const next = structuredClone(current);
  next.metadata.big = pseudoRandomBase64(80 * 1024);

  const patch = diffDocumentStates(current, next);
  const commit = await storeA.createCommit({
    docId,
    parentCommitId: mainA.headCommitId,
    mergeParentCommitId: null,
    createdBy: actor.userId,
    createdAt: Date.now(),
    message: "big",
    patch,
    nextState: next,
  });

  await storeA.updateBranchHead(docId, "main", commit.id);

  await waitForCondition(async () => {
    const remoteCommit = await storeB.getCommit(commit.id);
    if (!remoteCommit) return false;
    const state = await storeB.getDocumentStateAtCommit(commit.id);
    return state.metadata.big === next.metadata.big;
  }, 20_000);

  const remoteState = await storeB.getDocumentStateAtCommit(commit.id);
  assert.deepEqual(remoteState, normalizeDocumentState(next));

  // A `1009` close would indicate the server rejected an outgoing update from providerA.
  assert.ok(!closeCodes.includes(1009), `Unexpected 1009 close codes: ${closeCodes.join(",")}`);
});

