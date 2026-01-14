import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";

test("CollabSession direct setters enforce viewer role (no Yjs mutation)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(session.setCellValue("Sheet1:0:0", "hacked"), /Permission denied/);
  await assert.rejects(session.setCellFormula("Sheet1:0:0", "=HACK()"), /Permission denied/);

  assert.equal(session.cells.has("Sheet1:0:0"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setters reject unparseable cell keys when permissions are configured (viewer)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(session.setCellValue("bad-key", "hacked"), /Invalid cellKey/);
  await assert.rejects(session.setCellFormula("bad-key", "=HACK()"), /Invalid cellKey/);

  assert.equal(session.cells.has("bad-key"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setters still parse supported legacy key formats when permissions are configured (viewer)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  const before = Y.encodeStateAsUpdate(doc);

  // `${sheetId}:${row},${col}` legacy encoding.
  await assert.rejects(session.setCellValue("Sheet1:0,0", "hacked"), /Permission denied/);
  await assert.rejects(session.setCellFormula("Sheet1:0,0", "=HACK()"), /Permission denied/);

  // `r{row}c{col}` legacy encoding (resolved against defaultSheetId).
  await assert.rejects(session.setCellValue("r0c0", "hacked"), /Permission denied/);
  await assert.rejects(session.setCellFormula("r0c0", "=HACK()"), /Permission denied/);

  assert.equal(session.cells.has("Sheet1:0,0"), false);
  assert.equal(session.cells.has("r0c0"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setters enforce rangeRestrictions for editors (no Yjs mutation)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });

  session.setPermissions({
    role: "editor",
    userId: "u-editor",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        // Anyone can read, but only u-other can edit.
        readAllowlist: [],
        editAllowlist: ["u-other"],
      },
    ],
  });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(session.setCellValue("Sheet1:0:0", "hacked"), /Permission denied/);
  await assert.rejects(session.setCellFormula("Sheet1:0:0", "=HACK()"), /Permission denied/);

  assert.equal(session.cells.has("Sheet1:0:0"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setters reject unparseable cell keys when rangeRestrictions are configured (editor)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });

  session.setPermissions({
    role: "editor",
    userId: "u-editor",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        // Anyone can read, but only u-other can edit.
        readAllowlist: [],
        editAllowlist: ["u-other"],
      },
    ],
  });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(session.setCellValue("bad-key", "hacked"), /Invalid cellKey/);
  await assert.rejects(session.setCellFormula("bad-key", "=HACK()"), /Invalid cellKey/);

  assert.equal(session.cells.has("bad-key"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setters still parse supported legacy key formats when rangeRestrictions are configured (editor)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });

  session.setPermissions({
    role: "editor",
    userId: "u-editor",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        // Anyone can read, but only u-other can edit.
        readAllowlist: [],
        editAllowlist: ["u-other"],
      },
    ],
  });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(session.setCellValue("Sheet1:0,0", "hacked"), /Permission denied/);
  await assert.rejects(session.setCellFormula("Sheet1:0,0", "=HACK()"), /Permission denied/);
  await assert.rejects(session.setCellValue("r0c0", "hacked"), /Permission denied/);
  await assert.rejects(session.setCellFormula("r0c0", "=HACK()"), /Permission denied/);

  assert.equal(session.cells.has("Sheet1:0,0"), false);
  assert.equal(session.cells.has("r0c0"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setter escape hatch can bypass permissions when requested", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  await session.setCellValue("Sheet1:0:0", "allowed", { ignorePermissions: true });
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "allowed");

  await session.setCellFormula("Sheet1:0:1", "=1", { ignorePermissions: true });
  assert.equal((await session.getCell("Sheet1:0:1"))?.formula, "=1");

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setter escape hatch allows writing to unparseable cell keys (when encryption does not require parsing)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  await session.setCellValue("bad-key", "allowed", { ignorePermissions: true });
  assert.equal((await session.getCell("bad-key"))?.value, "allowed");

  session.destroy();
  doc.destroy();
});

test("CollabSession direct setter escape hatch still respects encryption invariants for unparseable cell keys", async () => {
  const docId = "collab-session-direct-setters-encryption-invalid-cellKey";
  const doc = new Y.Doc({ guid: docId });

  // Encryption enabled => setters require a parseable cell address so AAD can be bound to coordinates.
  const keyBytes = new Uint8Array(32).fill(7);
  const session = createCollabSession({
    doc,
    schema: { autoInit: false },
    encryption: { keyForCell: () => ({ keyId: "k1", keyBytes }) },
  });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(
    session.setCellValue("bad-key", "hacked", { ignorePermissions: true }),
    /expected "SheetId:row:col"/,
  );
  await assert.rejects(
    session.setCellFormula("bad-key", "=HACK()", { ignorePermissions: true }),
    /expected "SheetId:row:col"/,
  );

  assert.equal(session.cells.has("bad-key"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession permission getters expose role/capabilities and return defensive copies", () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });

  assert.equal(session.getPermissions(), null);
  assert.equal(session.getRole(), null);
  assert.equal(session.canComment(), false);
  assert.equal(session.canShare(), false);
  assert.equal(session.isReadOnly(), false);

  session.setPermissions({ role: "admin", userId: "u-admin", rangeRestrictions: [] });

  assert.deepEqual(session.getPermissions(), { role: "admin", userId: "u-admin", rangeRestrictions: [] });
  assert.equal(session.getRole(), "admin");
  assert.equal(session.canComment(), true);
  assert.equal(session.canShare(), true);
  assert.equal(session.isReadOnly(), false);

  const perms = session.getPermissions();
  assert.ok(perms, "expected getPermissions() to return a value after setPermissions()");
  perms.rangeRestrictions.push({
    sheetName: "Sheet1",
    startRow: 0,
    startCol: 0,
    endRow: 0,
    endCol: 0,
    readAllowlist: [],
    editAllowlist: ["someone-else"],
  });

  // Mutating the returned array should not affect the session's internal restrictions.
  assert.equal(session.canEditCell({ sheetId: "Sheet1", row: 0, col: 0 }), true);

  session.destroy();
  doc.destroy();
});
