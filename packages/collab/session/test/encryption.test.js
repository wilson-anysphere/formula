import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";
import { encryptCellPlaintext } from "../../encryption/src/index.node.js";

const REMOTE_ORIGIN = Symbol("remote");

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

test("CollabSession E2E cell encryption: encrypted in Yjs, decrypted with key, masked without key (in-memory sync)", async () => {
  const docId = "collab-session-encryption-test-doc";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForProtected = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && (cell.col === 0 || cell.col === 1)) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const sessionA = createCollabSession({
    doc: docA,
    encryption: { keyForCell: keyForProtected },
  });
  const sessionB = createCollabSession({ doc: docB });

  // Simulate historical docs where the same cell is stored under a legacy key encoding.
  // When we later encrypt the cell, we must not leave plaintext behind under the legacy key.
  docA.transact(() => {
    const legacy = new Y.Map();
    legacy.set("value", "old-plaintext");
    sessionA.cells.set("Sheet1:0,0", legacy);
  });

  await sessionA.setCellValue("Sheet1:0:0", "top-secret");

  // Raw Yjs should not contain plaintext.
  const cellMap = sessionA.cells.get("Sheet1:0:0");
  assert.ok(cellMap, "expected Yjs cell map to exist");
  assert.equal(typeof cellMap.get, "function");
  assert.equal(cellMap.get("value"), undefined);
  assert.equal(cellMap.get("formula"), undefined);
  assert.ok(cellMap.get("enc"), "expected encrypted payload under `enc`");
  assert.equal(JSON.stringify(cellMap.toJSON()).includes("top-secret"), false);

  const legacyCellMap = sessionA.cells.get("Sheet1:0,0");
  assert.ok(legacyCellMap, "expected legacy key cell map to exist");
  assert.equal(legacyCellMap.get("value"), undefined);
  assert.equal(legacyCellMap.get("formula"), undefined);
  assert.ok(legacyCellMap.get("enc"), "expected encrypted payload to overwrite legacy key value");
  assert.equal(JSON.stringify(legacyCellMap.toJSON()).includes("old-plaintext"), false);

  assert.equal((await sessionA.getCell("Sheet1:0:0"))?.value, "top-secret");

  const masked = await sessionB.getCell("Sheet1:0:0");
  assert.equal(masked?.value, "###");
  assert.equal(masked?.formula, null);
  assert.equal(masked?.encrypted, true);

  // Permission-like helpers should reflect that encrypted cells are unreadable/uneditable
  // without the relevant encryption key.
  assert.equal(sessionB.canReadCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);
  assert.equal(sessionB.canEditCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);

  // safeSet* APIs should fail gracefully (return false) rather than throwing when
  // the cell is encrypted but the key is unavailable.
  assert.equal(await sessionB.safeSetCellValue("Sheet1:0:0", "hacked"), false);
  assert.equal(await sessionB.safeSetCellFormula("Sheet1:0:0", "=HACK()"), false);

  // Legacy key encodings should still be treated as encrypted for read/edit gating.
  await sessionA.setCellValue("Sheet1:0,1", "legacy-secret");
  const maskedLegacy = await sessionB.getCell("Sheet1:0,1");
  assert.equal(maskedLegacy?.value, "###");
  assert.equal(maskedLegacy?.encrypted, true);

  // Canonical callers should still observe legacy-stored encrypted cells.
  const maskedLegacyCanonical = await sessionB.getCell("Sheet1:0:1");
  assert.equal(maskedLegacyCanonical?.value, "###");
  assert.equal(maskedLegacyCanonical?.encrypted, true);
  assert.equal(sessionB.canReadCell({ sheetId: "Sheet1", row: 0, col: 1 }), false);
  assert.equal(sessionB.canEditCell({ sheetId: "Sheet1", row: 0, col: 1 }), false);
  assert.equal(await sessionB.safeSetCellValue("Sheet1:0:1", "hacked-legacy"), false);
  assert.equal(sessionB.cells.has("Sheet1:0:1"), false);
  await assert.rejects(sessionB.setCellValue("Sheet1:0:1", "hacked-legacy-direct"));

  // Now "grant" the key by recreating the session with a resolver.
  sessionB.destroy();
  const sessionBWithKey = createCollabSession({ doc: docB, encryption: { keyForCell: keyForProtected } });
  assert.equal((await sessionBWithKey.getCell("Sheet1:0:0"))?.value, "top-secret");
  assert.equal((await sessionBWithKey.getCell("Sheet1:0,1"))?.value, "legacy-secret");
  assert.equal((await sessionBWithKey.getCell("Sheet1:0:1"))?.value, "legacy-secret");

  sessionA.destroy();
  sessionBWithKey.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession refuses to overwrite encrypted cells when keyId does not match the enc payload", async () => {
  const docId = "collab-session-encryption-keyid-mismatch-test-doc";
  const doc = new Y.Doc({ guid: docId });

  const keyBytesA = new Uint8Array(32).fill(7);
  const keyBytesB = new Uint8Array(32).fill(8);

  const keyForA1Correct = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes: keyBytesA };
    }
    return null;
  };
  const keyForA1Wrong = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-2", keyBytes: keyBytesB };
    }
    return null;
  };

  const sessionA = createCollabSession({ doc, encryption: { keyForCell: keyForA1Correct } });
  await sessionA.setCellValue("Sheet1:0:0", "top-secret");
  const encBefore = sessionA.cells.get("Sheet1:0:0").get("enc");

  const sessionB = createCollabSession({ doc, encryption: { keyForCell: keyForA1Wrong } });

  // Permission-like helpers should reflect that a mismatched key cannot edit the cell.
  assert.equal(sessionB.canEditCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);
  assert.equal(await sessionB.safeSetCellValue("Sheet1:0:0", "hacked"), false);
  assert.equal(sessionA.cells.get("Sheet1:0:0").get("enc"), encBefore);

  // Direct setters should also refuse to clobber the ciphertext.
  await assert.rejects(sessionB.setCellValue("Sheet1:0:0", "hacked-direct"), /key id mismatch|mismatch/i);
  assert.equal(sessionA.cells.get("Sheet1:0:0").get("enc"), encBefore);

  sessionA.destroy();
  sessionB.destroy();
  doc.destroy();
});

test("CollabSession refuses to overwrite cells when the enc payload schema is unsupported", async () => {
  const docId = "collab-session-encryption-unsupported-payload-test-doc";
  const doc = new Y.Doc({ guid: docId });

  // Simulate a future encryption payload version that this client cannot parse.
  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", {
      v: 2,
      alg: "AES-256-GCM",
      keyId: "k-range-1",
      ivBase64: "AA==",
      tagBase64: "AA==",
      ciphertextBase64: "AA==",
    });
    doc.getMap("cells").set("Sheet1:0:0", cell);
  });

  const keyBytes = new Uint8Array(32).fill(7);
  const session = createCollabSession({
    doc,
    encryption: {
      keyForCell: () => ({ keyId: "k-range-1", keyBytes }),
    },
  });

  // Unknown schemas are treated as non-editable to avoid clobbering newer ciphertext.
  assert.equal(session.canEditCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);
  assert.equal(await session.safeSetCellValue("Sheet1:0:0", "hacked"), false);
  await assert.rejects(session.setCellValue("Sheet1:0:0", "hacked-direct"), /unsupported encrypted cell payload/i);

  const encAfter = doc.getMap("cells").get("Sheet1:0:0").get("enc");
  assert.equal(encAfter.v, 2);

  session.destroy();
  doc.destroy();
});

test("CollabSession setCells refuses to write plaintext when an enc marker is present under a non-canonical key (null enc payload)", async () => {
  const docId = "collab-session-encryption-null-enc-marker-test-doc";
  const doc = new Y.Doc({ guid: docId });

  // Simulate a legacy/foreign doc that stores an explicit `enc: null` marker under
  // a non-canonical key encoding (empty sheetId resolves to defaultSheetId).
  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("enc", null);
    doc.getMap("cells").set(":0:0", cell);
  });

  const session = createCollabSession({ doc });

  await assert.rejects(
    session.setCells([{ cellKey: ":0:0", value: "leak" }]),
    /Missing encryption key/i,
  );

  const raw = doc.getMap("cells").get(":0:0");
  assert.ok(raw, "expected Yjs cell map to remain present");
  assert.equal(raw.get("enc"), null);
  assert.equal(raw.get("value"), undefined);
  assert.equal(raw.get("formula"), undefined);

  session.destroy();
  doc.destroy();
});

test("CollabSession prefers a non-null enc payload over an enc=null marker across legacy cell keys", async () => {
  const docId = "collab-session-encryption-null-marker-precedence-test-doc";
  const doc = new Y.Doc({ guid: docId });

  const keyBytes = new Uint8Array(32).fill(7);
  const keyId = "k-range-1";

  const enc = await encryptCellPlaintext({
    plaintext: { value: "top-secret", formula: null },
    key: { keyId, keyBytes },
    context: { docId, sheetId: "Sheet1", row: 0, col: 0 },
  });

  // Simulate a foreign/legacy doc state:
  // - canonical key has an `enc: null` marker (encryption marker)
  // - legacy key stores the real ciphertext payload
  doc.transact(() => {
    const cells = doc.getMap("cells");

    const marker = new Y.Map();
    marker.set("enc", null);
    cells.set("Sheet1:0:0", marker);

    const payload = new Y.Map();
    payload.set("enc", enc);
    cells.set("Sheet1:0,0", payload);
  });

  const session = createCollabSession({
    doc,
    encryption: {
      keyForCell: (cell) => {
        if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
          return { keyId, keyBytes };
        }
        return null;
      },
    },
  });

  assert.equal(session.canReadCell({ sheetId: "Sheet1", row: 0, col: 0 }), true);
  assert.equal(session.canEditCell({ sheetId: "Sheet1", row: 0, col: 0 }), true);

  const cell = await session.getCell("Sheet1:0:0");
  assert.equal(cell?.value ?? null, "top-secret");
  assert.equal(cell?.formula ?? null, null);

  // Ensure writes don't fail closed due to the marker when ciphertext exists on another alias.
  await session.setCellValue("Sheet1:0:0", "updated");
  const updated = await session.getCell("Sheet1:0:0");
  assert.equal(updated?.value ?? null, "updated");

  session.destroy();
  doc.destroy();
});

test("CollabSession E2E cell encryption: encryptFormat encrypts per-cell format and removes plaintext `format`", async () => {
  const docId = "collab-session-encryption-test-doc-encryptFormat";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const sessionA = createCollabSession({
    doc: docA,
    encryption: { keyForCell: keyForA1, encryptFormat: true },
  });
  const sessionB = createCollabSession({
    doc: docB,
    encryption: { keyForCell: () => null, encryptFormat: true },
  });

  const style = { font: { bold: true } };

  // Seed a plaintext format and then encrypt via `setCellValue`. The encrypted write should
  // migrate `format` into the ciphertext and remove it from the shared Yjs map.
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "plaintext");
    cell.set("format", style);
    sessionA.cells.set("Sheet1:0:0", cell);
  });

  await sessionA.setCellValue("Sheet1:0:0", "top-secret");

  const raw = sessionA.cells.get("Sheet1:0:0");
  assert.ok(raw, "expected Yjs cell map to exist");
  assert.equal(raw.get("format"), undefined);
  assert.equal(raw.get("style"), undefined);
  assert.ok(raw.get("enc"), "expected encrypted payload under `enc`");

  const decrypted = await sessionA.getCell("Sheet1:0:0");
  assert.deepEqual(decrypted?.format, style);

  const masked = await sessionB.getCell("Sheet1:0:0");
  assert.equal(masked?.value, "###");
  assert.equal(masked?.encrypted, true);
  assert.deepEqual(masked?.format, null);

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession E2E cell encryption: encryptFormat migrates legacy `style` key into ciphertext and removes plaintext", async () => {
  const docId = "collab-session-encryption-test-doc-encryptFormat-style";
  const docA = new Y.Doc({ guid: docId });
  const docB = new Y.Doc({ guid: docId });
  const disconnect = connectDocs(docA, docB);

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  const sessionA = createCollabSession({
    doc: docA,
    encryption: { keyForCell: keyForA1, encryptFormat: true },
  });

  const style = { font: { italic: true } };

  // Seed a legacy plaintext `style` field and then encrypt via `setCellValue`. The encrypted write should
  // migrate `style` into the ciphertext and remove it from the shared Yjs map.
  docA.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "plaintext");
    cell.set("style", style);
    sessionA.cells.set("Sheet1:0:0", cell);
  });

  await sessionA.setCellValue("Sheet1:0:0", "top-secret");

  const raw = sessionA.cells.get("Sheet1:0:0");
  assert.ok(raw, "expected Yjs cell map to exist");
  assert.equal(raw.get("format"), undefined);
  assert.equal(raw.get("style"), undefined);
  assert.ok(raw.get("enc"), "expected encrypted payload under `enc`");

  const decrypted = await sessionA.getCell("Sheet1:0:0");
  assert.deepEqual(decrypted?.format, style);

  sessionA.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession encryptFormat does not fall back to plaintext `format` for legacy encrypted payloads", async () => {
  const docId = "collab-session-encryption-test-doc-encryptFormat-no-plaintext-fallback";
  const doc = new Y.Doc({ guid: docId });

  const keyBytes = new Uint8Array(32).fill(7);
  const keyForA1 = (cell) => {
    if (cell.sheetId === "Sheet1" && cell.row === 0 && cell.col === 0) {
      return { keyId: "k-range-1", keyBytes };
    }
    return null;
  };

  // Simulate an older client that encrypts only { value, formula } but leaves plaintext `format`.
  const legacyWriter = createCollabSession({ doc, encryption: { keyForCell: keyForA1 } });
  await legacyWriter.setCellValue("Sheet1:0:0", "top-secret");
  doc.transact(() => {
    const cell = legacyWriter.cells.get("Sheet1:0:0");
    cell.set("format", { font: { bold: true } });
  });
  legacyWriter.destroy();

  const reader = createCollabSession({ doc, encryption: { keyForCell: keyForA1, encryptFormat: true } });
  const cell = await reader.getCell("Sheet1:0:0");
  assert.equal(cell?.value, "top-secret");
  // Confidentiality-first: do not apply plaintext formatting when `enc` is present.
  assert.deepEqual(cell?.format, null);

  reader.destroy();
  doc.destroy();
});
