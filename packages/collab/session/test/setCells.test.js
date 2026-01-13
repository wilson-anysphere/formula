import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";

test("CollabSession setCells writes the same Yjs shape as individual setCell* calls", async () => {
  const originalNow = Date.now;
  Date.now = () => 1_700_000_000_000;
  try {
    const docA = new Y.Doc();
    const docB = new Y.Doc();

    const sessionA = createCollabSession({ doc: docA });
    const sessionB = createCollabSession({ doc: docB });

    await sessionA.setCellValue("Sheet1:0:0", 123);
    await sessionA.setCellFormula("Sheet1:0:1", "=A1+1");
    await sessionA.setCellValue("Sheet1:1:0", "x");
    await sessionA.setCellFormula("Sheet1:1:1", null);
    // Formula wins when both formula+value are provided.
    await sessionA.setCellFormula("Sheet1:2:0", "=1");

    await sessionB.setCells([
      { cellKey: "Sheet1:0:0", value: 123 },
      { cellKey: "Sheet1:0:1", formula: "=A1+1" },
      { cellKey: "Sheet1:1:0", value: "x" },
      { cellKey: "Sheet1:1:1", formula: null },
      { cellKey: "Sheet1:2:0", value: "ignored", formula: "=1" },
    ]);

    assert.deepEqual(sessionB.cells.toJSON(), sessionA.cells.toJSON());

    sessionA.destroy();
    sessionB.destroy();
    docA.destroy();
    docB.destroy();
  } finally {
    Date.now = originalNow;
  }
});

test("CollabSession setCells rejects entire batch when any cell is not editable (permissions) and leaves doc unchanged", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });

  session.setPermissions({
    role: "editor",
    userId: "u-editor",
    rangeRestrictions: [
      {
        sheetName: "Sheet1",
        startRow: 0,
        startCol: 1,
        endRow: 0,
        endCol: 1,
        readAllowlist: [],
        // Only some other user can edit B1.
        editAllowlist: ["u-other"],
      },
    ],
  });

  const before = Y.encodeStateAsUpdate(doc);

  await assert.rejects(
    session.setCells([
      { cellKey: "Sheet1:0:0", value: "allowed" },
      { cellKey: "Sheet1:0:1", value: "blocked" },
    ]),
    (err) => String(err?.message ?? err).includes("Sheet1:0:1"),
  );

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  assert.deepEqual(session.cells.toJSON(), {});
  assert.equal(session.cells.has("Sheet1:0:0"), false);
  assert.equal(session.cells.has("Sheet1:0:1"), false);

  session.destroy();
  doc.destroy();
});

test("CollabSession setCells enforces viewer role permissions (no Yjs mutation)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  const before = Y.encodeStateAsUpdate(doc);
  await assert.rejects(
    session.setCells([{ cellKey: "Sheet1:0:0", value: "hacked" }]),
    /Permission denied/,
  );
  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);
  assert.equal(session.cells.has("Sheet1:0:0"), false);

  session.destroy();
  doc.destroy();
});

test("CollabSession setCells ignorePermissions bypasses permission checks (but still respects encryption invariants)", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  await session.setCells([{ cellKey: "Sheet1:0:0", value: "allowed" }], { ignorePermissions: true });
  assert.equal((await session.getCell("Sheet1:0:0"))?.value, "allowed");

  session.destroy();
  doc.destroy();
});

test("CollabSession setCells only bypasses permissions when ignorePermissions is explicitly true", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, schema: { autoInit: false } });
  session.setPermissions({ role: "viewer", userId: "u-viewer", rangeRestrictions: [] });

  const before = Y.encodeStateAsUpdate(doc);

  // `ignorePermissions` is an explicit escape hatch; non-boolean truthy values
  // should not bypass permission enforcement.
  await assert.rejects(
    // @ts-expect-error intentionally invalid type
    session.setCells([{ cellKey: "Sheet1:0:0", value: "hacked" }], { ignorePermissions: "true" }),
    /Permission denied/,
  );

  assert.equal(session.cells.has("Sheet1:0:0"), false);

  const after = Y.encodeStateAsUpdate(doc);
  assert.equal(Buffer.from(before).equals(Buffer.from(after)), true);

  session.destroy();
  doc.destroy();
});

test("CollabSession setCells encrypts protected cells and never writes plaintext into `enc` cells", async () => {
  const docId = "collab-session-setCells-encryption-test-doc";
  const doc = new Y.Doc({ guid: docId });

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForProtected = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const session = createCollabSession({ doc, encryption: { keyForCell: keyForProtected } });

  await session.setCells([
    { cellKey: "Sheet1:0:0", value: "top-secret" },
    { cellKey: "Sheet1:0:1", value: "plaintext" },
  ]);

  const encCell = session.cells.get("Sheet1:0:0");
  assert.ok(encCell, "expected encrypted cell map to exist");
  assert.equal(encCell.get("value"), undefined);
  assert.equal(encCell.get("formula"), undefined);
  assert.ok(encCell.get("enc"), "expected encrypted payload under `enc`");
  assert.equal(JSON.stringify(encCell.toJSON()).includes("top-secret"), false);

  const plainCell = session.cells.get("Sheet1:0:1");
  assert.ok(plainCell, "expected plaintext cell map to exist");
  assert.equal(plainCell.get("enc"), undefined);
  assert.equal(plainCell.get("value"), "plaintext");
  assert.equal(plainCell.get("formula"), null);

  session.destroy();
  doc.destroy();
});

test("CollabSession setCells encryptFormat migrates plaintext `format` into the encrypted payload and removes plaintext", async () => {
  const docId = "collab-session-setCells-encryption-test-doc-encryptFormat";
  const doc = new Y.Doc({ guid: docId });

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForProtected = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const session = createCollabSession({ doc, encryption: { keyForCell: keyForProtected, encryptFormat: true } });
  const sessionNoKey = createCollabSession({ doc, encryption: { keyForCell: () => null, encryptFormat: true } });

  const style = { font: { bold: true } };
  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "plaintext");
    cell.set("format", style);
    session.cells.set("Sheet1:0:0", cell);
  });

  await session.setCells([{ cellKey: "Sheet1:0:0", value: "top-secret" }]);

  const encCell = session.cells.get("Sheet1:0:0");
  assert.ok(encCell, "expected encrypted cell map to exist");
  assert.equal(encCell.get("value"), undefined);
  assert.equal(encCell.get("formula"), undefined);
  assert.equal(encCell.get("format"), undefined);
  assert.equal(encCell.get("style"), undefined);
  assert.ok(encCell.get("enc"), "expected encrypted payload under `enc`");

  const decrypted = await session.getCell("Sheet1:0:0");
  assert.deepEqual(decrypted?.format, style);

  const masked = await sessionNoKey.getCell("Sheet1:0:0");
  assert.equal(masked?.value, "###");
  assert.equal(masked?.encrypted, true);
  assert.deepEqual(masked?.format, null);

  session.destroy();
  sessionNoKey.destroy();
  doc.destroy();
});
