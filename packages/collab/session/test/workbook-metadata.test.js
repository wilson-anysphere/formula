import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";
import {
  createMetadataManagerForSession,
  createNamedRangeManagerForSession,
  createSheetManagerForSession,
} from "@formula/collab-workbook";

import { createCollabSession } from "../src/index.ts";

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function connectDocs(docA, docB) {
  const forwardA = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docB, update, REMOTE_ORIGIN);
  };
  const forwardB = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docA, update, REMOTE_ORIGIN);
  };

  docA.on("update", forwardA);
  docB.on("update", forwardB);

  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

/**
 * @param {import("../src/index.ts").CollabSession} session
 */
function snapshotSheets(session) {
  return session.sheets.toArray().map((sheet) => ({
    id: String(sheet.get("id") ?? ""),
    name: sheet.get("name") == null ? null : String(sheet.get("name")),
  }));
}

test("CollabSession workbook metadata: sheets + namedRanges sync and local undo (in-memory)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });

  const sheetsA = createSheetManagerForSession(sessionA);
  const namedRangesA = createNamedRangeManagerForSession(sessionA);
  const metadataA = createMetadataManagerForSession(sessionA);

  // Default schema initialization should ensure at least one sheet.
  assert.equal(sessionA.sheets.length >= 1, true);
  assert.deepEqual(snapshotSheets(sessionA), snapshotSheets(sessionB));

  sheetsA.addSheet({ id: "Sheet2", name: "Sheet2" });
  sessionA.undo?.stopCapturing();
  assert.deepEqual(snapshotSheets(sessionB).map((s) => s.id), ["Sheet1", "Sheet2"]);

  sheetsA.renameSheet("Sheet2", "Budget");
  sessionA.undo?.stopCapturing();
  assert.deepEqual(snapshotSheets(sessionB).find((s) => s.id === "Sheet2")?.name, "Budget");

  // Move forward: sheet1 from index 0 -> 1 should result in ["Sheet2", "Sheet1"].
  sheetsA.moveSheet("Sheet1", 1);
  sessionA.undo?.stopCapturing();
  assert.deepEqual(snapshotSheets(sessionB).map((s) => s.id), ["Sheet2", "Sheet1"]);

  namedRangesA.set("MyRange", { sheetId: "Sheet2", range: "A1:B2" });
  sessionA.undo?.stopCapturing();
  assert.deepEqual(sessionB.namedRanges.get("MyRange"), { sheetId: "Sheet2", range: "A1:B2" });

  namedRangesA.set("MyRange", { sheetId: "Sheet2", range: "A1:C3" });
  sessionA.undo?.stopCapturing();
  assert.deepEqual(sessionB.namedRanges.get("MyRange"), { sheetId: "Sheet2", range: "A1:C3" });

  metadataA.set("title", "Quarterly Budget");
  sessionA.undo?.stopCapturing();
  assert.equal(sessionB.metadata.get("title"), "Quarterly Budget");

  // Undo should only revert local-origin changes (and sync that undo update to peers).
  sessionA.undo?.undo();
  assert.equal(sessionA.metadata.has("title"), false);
  assert.equal(sessionB.metadata.has("title"), false);

  sessionA.undo?.undo();
  assert.deepEqual(sessionA.namedRanges.get("MyRange"), { sheetId: "Sheet2", range: "A1:B2" });
  assert.deepEqual(sessionB.namedRanges.get("MyRange"), { sheetId: "Sheet2", range: "A1:B2" });

  sessionA.undo?.undo();
  assert.equal(sessionA.namedRanges.has("MyRange"), false);
  assert.equal(sessionB.namedRanges.has("MyRange"), false);

  sessionA.undo?.undo();
  assert.deepEqual(snapshotSheets(sessionA).map((s) => s.id), ["Sheet1", "Sheet2"]);
  assert.deepEqual(snapshotSheets(sessionB).map((s) => s.id), ["Sheet1", "Sheet2"]);

  sessionA.undo?.undo();
  assert.equal(snapshotSheets(sessionA).find((s) => s.id === "Sheet2")?.name, "Sheet2");
  assert.equal(snapshotSheets(sessionB).find((s) => s.id === "Sheet2")?.name, "Sheet2");

  sessionA.undo?.undo();
  assert.deepEqual(snapshotSheets(sessionA).map((s) => s.id), ["Sheet1"]);
  assert.deepEqual(snapshotSheets(sessionB).map((s) => s.id), ["Sheet1"]);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession schema normalizes duplicate sheet ids when docs merge (in-memory)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  // Each session initializes its local schema independently (offline).
  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });

  assert.deepEqual(snapshotSheets(sessionA).map((s) => s.id), ["Sheet1"]);
  assert.deepEqual(snapshotSheets(sessionB).map((s) => s.id), ["Sheet1"]);

  // When the docs are later connected, we should not end up with duplicate Sheet1 entries.
  const disconnect = connectDocs(docA, docB);
  assert.deepEqual(snapshotSheets(sessionA).map((s) => s.id), ["Sheet1"]);
  assert.deepEqual(snapshotSheets(sessionB).map((s) => s.id), ["Sheet1"]);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession schema does not delete sheets that are populated across multiple transactions", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sessionA = createCollabSession({ doc: docA });
  const sessionB = createCollabSession({ doc: docB });
  const disconnect = connectDocs(docA, docB);

  // Simulate a (sloppy but possible) sheet creation flow where the sheet map is
  // inserted into the array and populated in a later transaction.
  const sheetsB = sessionB.sheets;
  const sheet = new Y.Map();
  docB.transact(() => {
    sheetsB.push([sheet]);
  });

  docB.transact(() => {
    sheet.set("id", "Sheet2");
    sheet.set("name", "Sheet2");
  });

  const sheetIdsA = snapshotSheets(sessionA).map((s) => s.id);
  const sheetIdsB = snapshotSheets(sessionB).map((s) => s.id);

  assert.ok(sheetIdsA.includes("Sheet2"));
  assert.ok(sheetIdsB.includes("Sheet2"));

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession schema assigns default id/name when sheets contain an entry without an id", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");

  // Malformed sheet entry (missing id/name). The schema initializer should
  // salvage this by assigning it the default sheet id/name.
  sheets.push([new Y.Map()]);

  const session = createCollabSession({ doc });

  assert.equal(session.sheets.length, 1);
  const sheet = session.sheets.get(0);
  assert.ok(sheet);
  assert.equal(sheet.get("id"), "Sheet1");
  assert.equal(sheet.get("name"), "Sheet1");

  session.destroy();
  doc.destroy();
});

test("CollabSession workbook metadata undo never reverts other users' overwrites (in-memory)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });

  const sheetsA = createSheetManagerForSession(sessionA);
  const sheetsB = createSheetManagerForSession(sessionB);
  const namedRangesA = createNamedRangeManagerForSession(sessionA);
  const namedRangesB = createNamedRangeManagerForSession(sessionB);
  const metadataA = createMetadataManagerForSession(sessionA);
  const metadataB = createMetadataManagerForSession(sessionB);

  // Setup a shared sheet and then have B overwrite A's rename. Undoing on A should
  // *not* undo B's rename.
  sheetsA.addSheet({ id: "Sheet2", name: "Sheet2" });
  sessionA.undo?.stopCapturing();

  sheetsA.renameSheet("Sheet2", "Budget A");
  sessionA.undo?.stopCapturing();

  sheetsB.renameSheet("Sheet2", "Budget B");
  sessionB.undo?.stopCapturing();

  assert.equal(snapshotSheets(sessionA).find((s) => s.id === "Sheet2")?.name, "Budget B");
  sessionA.undo?.undo();
  assert.equal(snapshotSheets(sessionA).find((s) => s.id === "Sheet2")?.name, "Budget B");
  assert.equal(snapshotSheets(sessionB).find((s) => s.id === "Sheet2")?.name, "Budget B");

  // Named range overwrites.
  namedRangesA.set("NR1", { range: "A1:B2" });
  sessionA.undo?.stopCapturing();
  namedRangesB.set("NR1", { range: "C1:D2" });
  sessionB.undo?.stopCapturing();

  assert.deepEqual(sessionA.namedRanges.get("NR1"), { range: "C1:D2" });
  sessionA.undo?.undo();
  assert.deepEqual(sessionA.namedRanges.get("NR1"), { range: "C1:D2" });
  assert.deepEqual(sessionB.namedRanges.get("NR1"), { range: "C1:D2" });

  // Metadata overwrites.
  metadataA.set("title", "From A");
  sessionA.undo?.stopCapturing();
  metadataB.set("title", "From B");
  sessionB.undo?.stopCapturing();

  assert.equal(sessionA.metadata.get("title"), "From B");
  sessionA.undo?.undo();
  assert.equal(sessionA.metadata.get("title"), "From B");
  assert.equal(sessionB.metadata.get("title"), "From B");

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});
