import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { randomUUID } from "node:crypto";
import { createRequire } from "node:module";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { createUndoService } from "../packages/collab/undo/index.js";
import { maskCellValue } from "../packages/collab/permissions/index.js";
import { createCollabSession } from "../packages/collab/session/src/index.ts";
import { bindYjsToDocumentController } from "../packages/collab/binder/index.js";
import {
  getAvailablePort,
  startSyncServer,
  waitForCondition,
} from "../services/sync-server/test/test-helpers.ts";

async function waitForCell(documentController, sheetId, coord, expected) {
  await waitForCondition(() => {
    const cell = documentController.getCell(sheetId, coord);
    return (cell.value ?? null) === (expected.value ?? null) && (cell.formula ?? null) === (expected.formula ?? null);
  }, 10_000);
}

async function waitForCellFormat(documentController, sheetId, coord, expectedFormat) {
  await waitForCondition(() => {
    const cell = documentController.getCell(sheetId, coord);
    const style =
      typeof documentController.getCellFormat === "function"
        ? documentController.getCellFormat(sheetId, coord)
        : documentController.styleTable.get(cell.styleId);
    if (expectedFormat == null) {
      return !style || (typeof style === "object" && Object.keys(style).length === 0);
    }
    try {
      assert.deepEqual(style, expectedFormat);
      return true;
    } catch {
      return false;
    }
  }, 30_000);
}

test("sync-server + collab-session + Yjs↔DocumentController binder: sync, undo, persistence", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-e2e-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  // Resolve deps from the sync-server package so this test doesn't need to add them to the root package.json.
  const requireFromSyncServer = createRequire(
    new URL("../services/sync-server/package.json", import.meta.url)
  );
  const WebSocket = requireFromSyncServer("ws");

  /** @type {Awaited<ReturnType<typeof startSyncServer>> | null} */
  let server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });
  t.after(async () => {
    await server?.stop();
  });

  const createClient = ({ wsUrl: wsBaseUrl, docId, token, user, activeSheet, encryption = null }) => {
    const session = createCollabSession({
      connection: {
        wsUrl: wsBaseUrl,
        docId,
        token,
        WebSocketPolyfill: WebSocket,
        disableBc: true,
      },
      presence: { user, activeSheet, throttleMs: 0 },
      defaultSheetId: activeSheet,
      encryption,
    });

    const ydoc = session.doc;
    const provider = session.provider;
    const undo = createUndoService({ mode: "collab", doc: ydoc, scope: session.cells });
    const documentController = new DocumentController();
    const binder = bindYjsToDocumentController({
      ydoc,
      documentController,
      undoService: undo,
      defaultSheetId: activeSheet,
      userId: user.id,
      encryption,
    });

    let destroyed = false;
    const destroy = () => {
      if (destroyed) return;
      destroyed = true;
      binder.destroy();
      session.destroy();
      ydoc.destroy();
    };

    return { ydoc, provider, session, undo, documentController, binder, destroy };
  };

  const docId = `e2e-${randomUUID()}`;

  const wsUrl = server.wsUrl;

  const secretKeyBytes = new Uint8Array(32).fill(7);
  const keyForSecret = (cell) => {
    // Encrypt D1 only (Sheet1 row 0 col 3).
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 3) {
      return { keyId: "k-range-1", keyBytes: secretKeyBytes };
    }
    return null;
  };

  const clientA = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-a", name: "User A", color: "#ff0000" },
    activeSheet: "Sheet1",
    encryption: { keyForCell: keyForSecret },
  });
  const clientB = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-b", name: "User B", color: "#00ff00" },
    activeSheet: "Sheet1",
  });

  t.after(() => {
    clientA.destroy();
    clientB.destroy();
  });

  await Promise.all([clientA.session.whenSynced(), clientB.session.whenSynced()]);

  // After initial hydration, remote updates must not trigger full scans of the Yjs
  // cells map (the binder should be O(changed-cells), not O(total-cells)).
  const cellsB = clientB.session.cells;
  const originalForEach = cellsB.forEach;
  cellsB.forEach = () => {
    throw new Error("binder performed a full cells-map scan after hydration");
  };
  t.after(() => {
    cellsB.forEach = originalForEach;
  });

  // --- Edits propagate A -> B ---
  await clientA.session.setCellFormula("Sheet1:0:1", "=1+1");
  clientA.documentController.setCellValue("Sheet1", "A1", "hello");

  await waitForCell(clientB.documentController, "Sheet1", "B1", { value: null, formula: "=1+1" });
  await waitForCell(clientB.documentController, "Sheet1", "A1", { value: "hello", formula: null });
  assert.equal((await clientB.session.getCell("Sheet1:0:1"))?.formula, "=1+1");
  assert.equal((await clientB.session.getCell("Sheet1:0:0"))?.value, "hello");

  // --- Edits propagate B -> A ---
  clientB.documentController.setCellValue("Sheet1", "C1", 123);
  await waitForCell(clientA.documentController, "Sheet1", "C1", { value: 123, formula: null });
  assert.equal((await clientA.session.getCell("Sheet1:0:2"))?.value, 123);

  // --- Undo only affects local-origin changes ---
  clientA.undo.undo();

  await waitForCell(clientA.documentController, "Sheet1", "A1", { value: null, formula: null });
  await waitForCell(clientB.documentController, "Sheet1", "A1", { value: null, formula: null });

  assert.equal(clientA.documentController.getCell("Sheet1", "B1").formula, "=1+1");
  assert.equal(clientB.documentController.getCell("Sheet1", "B1").formula, "=1+1");

  assert.equal(clientA.documentController.getCell("Sheet1", "C1").value, 123);
  assert.equal(clientB.documentController.getCell("Sheet1", "C1").value, 123);

  // --- Encrypted cells: authorized clients see plaintext, unauthorized clients see masked ---
  clientA.documentController.setCellValue("Sheet1", "D1", "top-secret");

  await waitForCell(clientB.documentController, "Sheet1", "D1", { value: "###", formula: null });
  assert.equal(clientA.documentController.getCell("Sheet1", "D1").value, "top-secret");

  {
    const cell = await clientB.session.getCell("Sheet1:0:3");
    assert.equal(cell?.value, "###");
    assert.equal(cell?.formula, null);
    assert.equal(cell?.encrypted, true);
  }

  // --- Sheet view state propagates + persists (freeze panes, row/col sizes) ---
  clientA.documentController.setFrozen("Sheet1", 2, 1);
  clientA.documentController.setColWidth("Sheet1", 0, 120);
  clientA.documentController.setRowHeight("Sheet1", 5, 40);

  await waitForCondition(() => {
    const view = clientB.documentController.getSheetView("Sheet1");
    return (
      view.frozenRows === 2 &&
      view.frozenCols === 1 &&
      view.colWidths?.["0"] === 120 &&
      view.rowHeights?.["5"] === 40
    );
  }, 10_000);

  // Now simulate a remote collaborator updating view state directly in Yjs
  // (i.e. not via DocumentController). This should propagate back into both
  // clients' DocumentControllers via the binder's `sheets.observeDeep`.
  const remoteViewOrigin = { type: "remote-view-test" };
  clientB.ydoc.transact(() => {
    const sheets = clientB.ydoc.getArray("sheets");
    const entry = sheets.get(0);
    assert.ok(entry, "expected Sheet1 sheet entry");
    if (typeof (entry).set === "function") {
      entry.set("view", { frozenRows: 0, frozenCols: 3, colWidths: { "1": 200 }, rowHeights: { "10": 55 } });
    } else {
      // Plain-object fallback (should be rare in sync-server scenarios).
      const idx = sheets.toArray().findIndex((e) => (e?.get?.("id") ?? e?.id) === "Sheet1");
      sheets.delete(idx, 1);
      sheets.insert(idx, [
        { id: "Sheet1", name: "Sheet1", view: { frozenRows: 0, frozenCols: 3, colWidths: { "1": 200 }, rowHeights: { "10": 55 } } },
      ]);
    }
  }, remoteViewOrigin);

  await waitForCondition(() => {
    const view = clientA.documentController.getSheetView("Sheet1");
    return (
      view.frozenRows === 0 &&
      view.frozenCols === 3 &&
      view.colWidths?.["1"] === 200 &&
      view.rowHeights?.["10"] === 55 &&
      !view.colWidths?.["0"] &&
      !view.rowHeights?.["5"]
    );
  }, 10_000);

  // Tear down clients and restart the server, keeping the same data directory.
  clientA.destroy();
  clientB.destroy();

  // Give the server a moment to persist state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 500));

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });

  const clientC = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-c", name: "User C", color: "#0000ff" },
    activeSheet: "Sheet1",
  });

  t.after(() => {
    clientC.destroy();
  });

  await clientC.session.whenSynced();

  // --- Hydration from persisted sync-server state ---
  await waitForCell(clientC.documentController, "Sheet1", "B1", { value: null, formula: "=1+1" });
  await waitForCell(clientC.documentController, "Sheet1", "C1", { value: 123, formula: null });
  await waitForCell(clientC.documentController, "Sheet1", "A1", { value: null, formula: null });
  await waitForCell(clientC.documentController, "Sheet1", "D1", { value: "###", formula: null });

  await waitForCondition(() => {
    const view = clientC.documentController.getSheetView("Sheet1");
    return (
      view.frozenRows === 0 &&
      view.frozenCols === 3 &&
      view.colWidths?.["1"] === 200 &&
      view.rowHeights?.["10"] === 55
    );
  }, 10_000);

  {
    const cell = await clientC.session.getCell("Sheet1:0:3");
    assert.equal(cell?.value, "###");
    assert.equal(cell?.formula, null);
    assert.equal(cell?.encrypted, true);
  }
});

test("sync-server + Yjs↔DocumentController binder: sync format-only cells", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-e2e-formatting-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  // Resolve deps from the sync-server package so this test doesn't need to add them to the root package.json.
  const requireFromSyncServer = createRequire(
    new URL("../services/sync-server/package.json", import.meta.url)
  );
  const WebSocket = requireFromSyncServer("ws");

  let server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });
  t.after(async () => {
    await server?.stop();
  });

  const createClient = ({ wsUrl: wsBaseUrl, docId, token, user, activeSheet }) => {
    const session = createCollabSession({
      connection: {
        wsUrl: wsBaseUrl,
        docId,
        token,
        WebSocketPolyfill: WebSocket,
        disableBc: true,
      },
      presence: { user, activeSheet, throttleMs: 0 },
      defaultSheetId: activeSheet,
    });

    const ydoc = session.doc;
    const undo = createUndoService({ mode: "collab", doc: ydoc, scope: session.cells });
    const documentController = new DocumentController();
    const binder = bindYjsToDocumentController({
      ydoc,
      documentController,
      undoService: undo,
      defaultSheetId: activeSheet,
      userId: user.id,
    });

    let destroyed = false;
    const destroy = () => {
      if (destroyed) return;
      destroyed = true;
      binder.destroy();
      session.destroy();
      ydoc.destroy();
    };

    return { session, undo, documentController, binder, destroy };
  };

  const docId = `e2e-formatting-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  const clientA = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-a", name: "User A", color: "#ff0000" },
    activeSheet: "Sheet1",
  });
  const clientB = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-b", name: "User B", color: "#00ff00" },
    activeSheet: "Sheet1",
  });

  t.after(() => {
    clientA.destroy();
    clientB.destroy();
  });

  await Promise.all([clientA.session.whenSynced(), clientB.session.whenSynced()]);

  // Apply formatting to the full height of column A (layered formatting) and ensure it propagates.
  clientA.documentController.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });
  await waitForCellFormat(clientB.documentController, "Sheet1", "A1", { font: { bold: true } });
  assert.equal(clientB.documentController.getCell("Sheet1", "A1").styleId, 0);

  // Ensure content+format updates round-trip.
  clientA.documentController.setCellValue("Sheet1", "A1", "hello");
  await waitForCell(clientB.documentController, "Sheet1", "A1", { value: "hello", formula: null });
  await waitForCellFormat(clientB.documentController, "Sheet1", "A1", { font: { bold: true } });
  assert.equal(clientB.documentController.getCell("Sheet1", "A1").styleId, 0);

  // Apply range-run formatting (large range, not a full column) and ensure it syncs + persists.
  clientA.documentController.setRangeFormat("Sheet1", "B1:B50001", { font: { italic: true } });
  await waitForCellFormat(clientB.documentController, "Sheet1", "B1", { font: { italic: true } });
  assert.equal(clientB.documentController.getCell("Sheet1", "B1").styleId, 0);

  // Tear down clients and restart the server to validate persistence.
  clientA.destroy();
  clientB.destroy();

  // Give the server a moment to persist state after the last client disconnects.
  await new Promise((r) => setTimeout(r, 500));

  await server.stop();
  server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });

  const clientC = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-c", name: "User C", color: "#0000ff" },
    activeSheet: "Sheet1",
  });

  t.after(() => {
    clientC.destroy();
  });

  await clientC.session.whenSynced();

  await waitForCell(clientC.documentController, "Sheet1", "A1", { value: "hello", formula: null });
  await waitForCellFormat(clientC.documentController, "Sheet1", "A1", { font: { bold: true } });
  assert.equal(clientC.documentController.getCell("Sheet1", "A1").styleId, 0);

  await waitForCellFormat(clientC.documentController, "Sheet1", "B1", { font: { italic: true } });
  assert.equal(clientC.documentController.getCell("Sheet1", "B1").styleId, 0);
});

test("sync-server + Yjs↔DocumentController binder: permission-masked cells are unreadable and uneditable", async (t) => {
  const dataDir = await mkdtemp(path.join(tmpdir(), "sync-server-e2e-permissions-"));
  t.after(async () => {
    await rm(dataDir, { recursive: true, force: true });
  });

  const port = await getAvailablePort();
  // Resolve deps from the sync-server package so this test doesn't need to add them to the root package.json.
  const requireFromSyncServer = createRequire(
    new URL("../services/sync-server/package.json", import.meta.url)
  );
  const WebSocket = requireFromSyncServer("ws");

  const server = await startSyncServer({
    port,
    dataDir,
    auth: { mode: "opaque", token: "test-token" },
  });
  t.after(async () => {
    await server.stop();
  });

  const restrictions = [
    // Lock A1 so only u-a can read/edit.
    {
      sheetName: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
      readAllowlist: ["u-a"],
      editAllowlist: ["u-a"],
    },
  ];

  const createClient = ({ wsUrl: wsBaseUrl, docId, token, user, activeSheet, permissions }) => {
    const session = createCollabSession({
      connection: {
        wsUrl: wsBaseUrl,
        docId,
        token,
        WebSocketPolyfill: WebSocket,
        disableBc: true,
      },
      presence: { user, activeSheet, throttleMs: 0 },
      defaultSheetId: activeSheet,
    });

    const ydoc = session.doc;
    const undo = createUndoService({ mode: "collab", doc: ydoc, scope: session.cells });
    const documentController = new DocumentController();
    const binder = bindYjsToDocumentController({
      ydoc,
      documentController,
      undoService: undo,
      defaultSheetId: activeSheet,
      userId: user.id,
      permissions,
    });

    let destroyed = false;
    const destroy = () => {
      if (destroyed) return;
      destroyed = true;
      binder.destroy();
      session.destroy();
      ydoc.destroy();
    };

    return { session, undo, documentController, binder, destroy };
  };

  const docId = `e2e-permissions-${randomUUID()}`;
  const wsUrl = server.wsUrl;

  const clientA = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-a", name: "User A", color: "#ff0000" },
    activeSheet: "Sheet1",
    permissions: { role: "editor", restrictions, userId: "u-a" },
  });
  const clientB = createClient({
    wsUrl,
    docId,
    token: "test-token",
    user: { id: "u-b", name: "User B", color: "#00ff00" },
    activeSheet: "Sheet1",
    permissions: { role: "editor", restrictions, userId: "u-b" },
  });

  t.after(() => {
    clientA.destroy();
    clientB.destroy();
  });

  await Promise.all([clientA.session.whenSynced(), clientB.session.whenSynced()]);

  // A writes a secret value; B should only see a masked placeholder.
  clientA.documentController.setCellValue("Sheet1", "A1", "top-secret");
  await waitForCell(clientA.documentController, "Sheet1", "A1", { value: "top-secret", formula: null });
  await waitForCell(clientB.documentController, "Sheet1", "A1", { value: maskCellValue("top-secret"), formula: null });

  // B attempts to edit the protected cell; binder should reject + revert locally and not push to Yjs.
  clientB.documentController.setCellValue("Sheet1", "A1", "hacked");
  await waitForCell(clientB.documentController, "Sheet1", "A1", { value: maskCellValue("top-secret"), formula: null });
  await waitForCell(clientA.documentController, "Sheet1", "A1", { value: "top-secret", formula: null });

  assert.equal((await clientA.session.getCell("Sheet1:0:0"))?.value, "top-secret");
  assert.equal((await clientB.session.getCell("Sheet1:0:0"))?.value, "top-secret");
});
