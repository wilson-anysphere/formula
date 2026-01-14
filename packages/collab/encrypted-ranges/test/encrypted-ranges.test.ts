import * as Y from "yjs";
import { describe, expect, it } from "vitest";
import { ensureWorkbookSchema } from "@formula/collab-workbook";

import { EncryptedRangeManager, createEncryptionPolicyFromDoc } from "../src/index.ts";

describe("@formula/collab-encrypted-ranges", () => {
  it("manager mutations can participate in an external collaborative UndoManager via custom transact", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const origin = { type: "test-encrypted-range-undo" };
    const undoManager = new Y.UndoManager(metadata, { trackedOrigins: new Set([origin]) });

    const mgr = new EncryptedRangeManager({ doc, transact: (fn) => doc.transact(fn, origin) });

    const id = mgr.add({ sheetId: "s1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });
    expect(mgr.list().map((r) => r.id)).toEqual([id]);
    expect(undoManager.canUndo()).toBe(true);

    undoManager.undo();
    expect(mgr.list()).toHaveLength(0);
    expect(undoManager.canRedo()).toBe(true);

    undoManager.redo();
    expect(mgr.list().map((r) => r.id)).toEqual([id]);
  });

  it("manager add refuses to overwrite unknown/corrupt encryptedRanges schemas", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const bogus: any = { foo: "bar" };
    doc.transact(() => {
      metadata.set("encryptedRanges", bogus);
    });

    const mgr = new EncryptedRangeManager({ doc });
    expect(() => mgr.list()).toThrow(/Unsupported metadata\.encryptedRanges schema/);
    expect(() =>
      mgr.add({ sheetId: "s1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" })
    ).toThrow(/Unsupported metadata\.encryptedRanges schema/);
    expect(() => mgr.update("r1", { endCol: 1 })).toThrow(/Unsupported metadata\.encryptedRanges schema/);
    expect(() => mgr.remove("r1")).toThrow(/Unsupported metadata\.encryptedRanges schema/);

    // Should not clobber the original value.
    expect(metadata.get("encryptedRanges")).toEqual(bogus);
  });

  it("policy helper fails closed when metadata.encryptedRanges is present in an unknown schema", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    doc.transact(() => {
      metadata.set("encryptedRanges", { foo: "bar" });
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe(null);

    // Still validates cell addresses.
    expect(policy.shouldEncryptCell({ sheetId: "", row: 0, col: 0 } as any)).toBe(false);
    expect(policy.keyIdForCell({ sheetId: "", row: 0, col: 0 } as any)).toBe(null);
  });

  it("manager add/list/remove is deterministic and dedupes identical ranges", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const mgr = new EncryptedRangeManager({ doc });

    const id1 = mgr.add({
      sheetId: " Sheet1 ",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
      keyId: "k1",
    });
    const id2 = mgr.add({
      sheetId: "Sheet1",
      startRow: 2,
      startCol: 0,
      endRow: 2,
      endCol: 0,
      keyId: "k1",
    });

    const list = mgr.list();
    expect(list.map((r) => r.id)).toEqual([id1, id2].sort());
    expect(list.find((r) => r.id === id1)?.sheetId).toBe("Sheet1");

    // Identical range should be deduped.
    const id1b = mgr.add({
      sheetId: "Sheet1",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
      keyId: "k1",
    });
    expect(id1b).toBe(id1);
    expect(mgr.list()).toHaveLength(2);

    mgr.remove(id1);
    expect(mgr.list().map((r) => r.id)).toEqual([id2]);
  });

  it("manager add resolves sheet display names to stable sheet ids when possible", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    // Sheet id differs from name.
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const mgr = new EncryptedRangeManager({ doc });

    mgr.add({ sheetId: "Budget", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });
    expect(mgr.list()).toMatchObject([{ sheetId: "sheet-123", keyId: "k1" }]);
  });

  it("manager add canonicalizes stable sheet id casing when possible", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "Sheet-ABC");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const mgr = new EncryptedRangeManager({ doc });
    mgr.add({ sheetId: "sheet-abc", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });

    expect(mgr.list()).toMatchObject([{ sheetId: "Sheet-ABC", keyId: "k1" }]);
  });

  it("manager update resolves sheet display names to stable sheet ids when possible", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const mgr = new EncryptedRangeManager({ doc });
    const id = mgr.add({ sheetId: "other-sheet", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });
    mgr.update(id, { sheetId: "Budget" });

    expect(mgr.list()).toMatchObject([{ id, sheetId: "sheet-123" }]);
  });

  it("manager update canonicalizes stable sheet id casing when possible", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "Sheet-ABC");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const mgr = new EncryptedRangeManager({ doc });
    const id = mgr.add({ sheetId: "other-sheet", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });
    mgr.update(id, { sheetId: "sheet-abc" });

    expect(mgr.list()).toMatchObject([{ id, sheetId: "Sheet-ABC" }]);
  });

  it("manager update validates and applies patches", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const mgr = new EncryptedRangeManager({ doc });
    const id = mgr.add({
      sheetId: "s1",
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 0,
      keyId: "k1",
    });

    mgr.update(id, { endRow: 1, endCol: 2, keyId: "k2" });
    expect(mgr.list().find((r) => r.id === id)).toEqual({
      id,
      sheetId: "s1",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 2,
      keyId: "k2",
    });

    // Invalid patch should throw and not modify the range.
    expect(() => mgr.update(id, { startRow: 10 })).toThrow(/startRow/);
    expect(mgr.list().find((r) => r.id === id)?.startRow).toBe(0);
  });

  it("manager validates inputs", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const mgr = new EncryptedRangeManager({ doc });

    expect(() =>
      mgr.add({
        sheetId: "",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      })
    ).toThrow(/sheetId/);

    expect(() =>
      mgr.add({
        sheetId: "s1",
        startRow: -1 as any,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      })
    ).toThrow(/startRow/);

    expect(() =>
      mgr.add({
        sheetId: "s1",
        startRow: 2,
        startCol: 0,
        endRow: 1,
        endCol: 0,
        keyId: "k1",
      })
    ).toThrow(/startRow.*endRow/);

    expect(() =>
      mgr.add({
        sheetId: "s1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "",
      })
    ).toThrow(/keyId/);
  });

  it("createEncryptionPolicyFromDoc shouldEncryptCell / keyIdForCell reflect metadata", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const mgr = new EncryptedRangeManager({ doc });
    const id = mgr.add({
      sheetId: "s1",
      startRow: 0,
      startCol: 0,
      endRow: 1,
      endCol: 1,
      keyId: "key-123",
    });

    const policy = createEncryptionPolicyFromDoc(doc);

    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(true);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 1, col: 1 })).toBe(true);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 2, col: 0 })).toBe(false);
    expect(policy.shouldEncryptCell({ sheetId: "other", row: 0, col: 0 })).toBe(false);

    expect(policy.keyIdForCell({ sheetId: "s1", row: 1, col: 1 })).toBe("key-123");
    expect(policy.keyIdForCell({ sheetId: "s1", row: 2, col: 0 })).toBe(null);

    mgr.remove(id);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(false);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe(null);
  });

  it("policy helper works when encryptedRanges are written directly to workbook metadata", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<Y.Map<unknown>>();
    const r = new Y.Map<unknown>();
    r.set("id", "r1");
    r.set("sheetId", "s1");
    r.set("startRow", 5);
    r.set("startCol", 2);
    r.set("endRow", 6);
    r.set("endCol", 4);
    r.set("keyId", "k1");
    ranges.push([r]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 5, col: 2 })).toBe(true);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 6, col: 4 })).toBe(true);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 4, col: 2 })).toBe(false);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 7, col: 2 })).toBe(false);
  });

  it("policy helper supports numeric fields stored as Y.Text", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    const r = new Y.Map<unknown>();
    r.set("id", "r1");
    r.set("sheetId", "s1");

    const startRow = new Y.Text();
    startRow.insert(0, "0");
    const startCol = new Y.Text();
    startCol.insert(0, "0");
    const endRow = new Y.Text();
    endRow.insert(0, "0");
    const endCol = new Y.Text();
    endCol.insert(0, "0");

    r.set("startRow", startRow);
    r.set("startCol", startCol);
    r.set("endRow", endRow);
    r.set("endCol", endCol);
    r.set("keyId", "k1");
    ranges.push([r]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe("k1");
  });

  it("policy helper supports legacy map schema (encryptedRanges as Y.Map<id, Y.Map>)", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Map<Y.Map<unknown>>();
    const r = new Y.Map<unknown>();
    // Intentionally omit `id` inside the value; use the map key as the id.
    r.set("sheetId", "s1");
    r.set("startRow", 0);
    r.set("startCol", 0);
    r.set("endRow", 0);
    r.set("endCol", 0);
    r.set("keyId", "k1");

    doc.transact(() => {
      ranges.set("range-1", r);
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe("k1");
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 1 })).toBe(false);
  });

  it("policy helper overlap precedence for legacy map schema prefers lexicographically greatest key", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Map<Y.Map<unknown>>();

    const r1 = new Y.Map<unknown>();
    r1.set("sheetId", "s1");
    r1.set("startRow", 0);
    r1.set("startCol", 0);
    r1.set("endRow", 0);
    r1.set("endCol", 0);
    r1.set("keyId", "k1");

    const r2 = new Y.Map<unknown>();
    r2.set("sheetId", "s1");
    r2.set("startRow", 0);
    r2.set("startCol", 0);
    r2.set("endRow", 0);
    r2.set("endCol", 0);
    r2.set("keyId", "k2");

    doc.transact(() => {
      ranges.set("a", r1);
      ranges.set("b", r2);
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe("k2");
  });

  it("manager can mutate legacy map schema by normalizing to the canonical array schema", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Map<Y.Map<unknown>>();
    const r = new Y.Map<unknown>();
    r.set("sheetId", "s1");
    r.set("startRow", 0);
    r.set("startCol", 0);
    r.set("endRow", 0);
    r.set("endCol", 0);
    r.set("keyId", "k1");

    doc.transact(() => {
      ranges.set("range-1", r);
      metadata.set("encryptedRanges", ranges);
    });

    const mgr = new EncryptedRangeManager({ doc });
    expect(mgr.list().map((e) => e.id)).toEqual(["range-1"]);

    // Update should work even though the original value omitted `id` (fallback is the map key).
    mgr.update("range-1", { endCol: 1 });
    expect(mgr.list().find((e) => e.id === "range-1")?.endCol).toBe(1);

    mgr.remove("range-1");
    expect(mgr.list()).toHaveLength(0);
  });

  it("policy helper supports legacy encryptedRanges entries without id", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    ranges.push([
      {
        // Legacy shape (pre-id): plain object entries in the encryptedRanges array.
        sheetId: "s1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      },
    ]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe("k1");
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 1 })).toBe(false);

    const mgr = new EncryptedRangeManager({ doc });
    const list = mgr.list();
    expect(list).toHaveLength(1);
    expect(list[0]!.id).toMatch(/^legacy:/);
  });

  it("manager update supports legacy Y.Map entries without id by persisting a derived id", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<Y.Map<unknown>>();
    const r = new Y.Map<unknown>();
    // Legacy-ish shape: a Y.Map entry without an explicit `id` field.
    r.set("sheetId", "s1");
    r.set("startRow", 0);
    r.set("startCol", 0);
    r.set("endRow", 0);
    r.set("endCol", 0);
    r.set("keyId", "k1");
    ranges.push([r]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const mgr = new EncryptedRangeManager({ doc });
    const [entry] = mgr.list();
    expect(entry?.id).toMatch(/^legacy:/);

    mgr.update(entry!.id, { endCol: 1 });
    const updated = mgr.list().find((e) => e.id === entry!.id);
    expect(updated?.endCol).toBe(1);

    // The update should persist the derived id into the underlying Y.Map so it remains stable
    // even though the legacy-id derivation would otherwise change when resizing the range.
    const stored = (metadata.get("encryptedRanges") as any).toArray?.()[0];
    expect(stored && typeof stored.get === "function").toBe(true);
    expect(stored.get("id")).toBe(entry!.id);
  });

  it("policy helper supports legacy `sheetName` field when `sheetId` is missing", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    ranges.push([
      {
        // Legacy-ish shape: `sheetName` instead of `sheetId`.
        sheetName: "s1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      },
    ]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe("k1");
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 1 })).toBe(false);
  });

  it("policy helper falls back to legacy `sheetName` when `sheetId` is present but blank", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    ranges.push([
      {
        // Corrupt/partial legacy shape: blank `sheetId` but a valid sheet name.
        sheetId: "   ",
        sheetName: "s1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      },
    ]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "s1", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe("k1");
  });

  it("policy helper resolves legacy sheet names to stable sheet ids via workbook sheets metadata", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    // Create a workbook sheet whose id differs from its display name.
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    ranges.push([
      {
        // Legacy shape: stores the sheet display name rather than the stable id.
        sheetName: "Budget",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      },
    ]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "sheet-123", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "sheet-123", row: 0, col: 0 })).toBe("k1");
  });

  it("policy helper matches stable sheet ids even when caller passes sheet display name", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    // Create a workbook sheet whose id differs from its display name.
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    // Store the encrypted range with the stable sheet id.
    const mgr = new EncryptedRangeManager({ doc });
    mgr.add({ sheetId: "sheet-123", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "Budget", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "Budget", row: 0, col: 0 })).toBe("k1");
    // Case-insensitive.
    expect(policy.keyIdForCell({ sheetId: "budget", row: 0, col: 0 })).toBe("k1");

    // If the sheet is renamed, the policy should keep matching by the *new* display name.
    doc.transact(() => {
      sheet.set("name", "Budget 2024");
    });
    expect(policy.shouldEncryptCell({ sheetId: "Budget 2024", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "Budget 2024", row: 0, col: 0 })).toBe("k1");

    // If the sheet entry is replaced (delete+insert with the same array length), the policy
    // should still resolve the display name for the stable id.
    const replaced = new Y.Map<unknown>();
    replaced.set("id", "sheet-123");
    replaced.set("name", "Budget 2025");
    doc.transact(() => {
      sheets.delete(0, 1);
      sheets.insert(0, [replaced]);
    });
    expect(policy.shouldEncryptCell({ sheetId: "Budget 2025", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "Budget 2025", row: 0, col: 0 })).toBe("k1");
  });

  it("policy helper does not treat other sheets' stable ids as the active sheet name", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    // Two sheets:
    // - sheet-123 has display name "Budget"
    // - a *different* sheet has stable id "Budget"
    const sheets = doc.getArray("sheets");
    const sheetA = new Y.Map<unknown>();
    sheetA.set("id", "sheet-123");
    sheetA.set("name", "Budget");
    const sheetB = new Y.Map<unknown>();
    sheetB.set("id", "Budget");
    sheetB.set("name", "Other");
    sheets.push([sheetA, sheetB]);

    const mgr = new EncryptedRangeManager({ doc });
    mgr.add({ sheetId: "Budget", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k-b" });

    const policy = createEncryptionPolicyFromDoc(doc);
    // Should match only the sheet with id "Budget".
    expect(policy.shouldEncryptCell({ sheetId: "Budget", row: 0, col: 0 })).toBe(true);
    // Should not match sheet-123 even though its display name is "Budget".
    expect(policy.shouldEncryptCell({ sheetId: "sheet-123", row: 0, col: 0 })).toBe(false);
  });

  it("policy helper matches legacy sheet names case-insensitively", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    // Create a workbook sheet whose id differs from its display name.
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    ranges.push([
      {
        // Legacy shape: stores the sheet display name (in different case) rather than the stable id.
        sheetName: "budget",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      },
    ]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "sheet-123", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "sheet-123", row: 0, col: 0 })).toBe("k1");
  });

  it("policy helper matches sheetId case-insensitively", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    ranges.push([
      {
        sheetId: "SHEET-123",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      },
    ]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "sheet-123", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "sheet-123", row: 0, col: 0 })).toBe("k1");
  });

  it("policy helper resolves legacy sheetName ranges even when the sheet id casing differs", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<any>();
    ranges.push([
      {
        sheetName: "Budget",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        keyId: "k1",
      },
    ]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.shouldEncryptCell({ sheetId: "SHEET-123", row: 0, col: 0 })).toBe(true);
    expect(policy.keyIdForCell({ sheetId: "SHEET-123", row: 0, col: 0 })).toBe("k1");
  });

  it("manager normalization rewrites legacy sheetName entries to stable sheet ids when possible", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<Y.Map<unknown>>();
    const r1 = new Y.Map<unknown>();
    r1.set("id", "r1");
    // Legacy shape: sheet display name, not sheet id.
    r1.set("sheetName", "Budget");
    r1.set("startRow", 0);
    r1.set("startCol", 0);
    r1.set("endRow", 0);
    r1.set("endCol", 0);
    r1.set("keyId", "k1");
    ranges.push([r1]);

    doc.transact(() => {
      metadata.set("encryptedRanges", ranges);
    });

    const mgr = new EncryptedRangeManager({ doc });
    // Trigger normalization via a tracked mutation.
    mgr.update("r1", { endCol: 1 });

    const storedRanges = metadata.get("encryptedRanges") as any;
    const stored = storedRanges?.get?.(0);
    expect(stored && typeof stored.get === "function").toBe(true);
    expect(stored.get("id")).toBe("r1");
    expect(stored.get("sheetId")).toBe("sheet-123");
    expect(stored.get("endCol")).toBe(1);
  });

  it("manager normalization dedupes legacy sheetName entries against canonical sheetId entries", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map<unknown>();
    sheet.set("id", "sheet-123");
    sheet.set("name", "Budget");
    sheets.push([sheet]);

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<Y.Map<unknown>>();

    const canonical = new Y.Map<unknown>();
    canonical.set("id", "a");
    canonical.set("sheetId", "sheet-123");
    canonical.set("startRow", 0);
    canonical.set("startCol", 0);
    canonical.set("endRow", 0);
    canonical.set("endCol", 0);
    canonical.set("keyId", "k1");

    const legacy = new Y.Map<unknown>();
    legacy.set("id", "b");
    // Same logical range, but stored with the sheet display name.
    legacy.set("sheetName", "Budget");
    legacy.set("startRow", 0);
    legacy.set("startCol", 0);
    legacy.set("endRow", 0);
    legacy.set("endCol", 0);
    legacy.set("keyId", "k1");

    doc.transact(() => {
      ranges.push([canonical, legacy]);
      metadata.set("encryptedRanges", ranges);
    });

    const mgr = new EncryptedRangeManager({ doc });
    // Mutate the legacy id; normalization should keep the mutated id and drop the duplicate.
    mgr.update("b", { endCol: 2 });

    const list = mgr.list();
    expect(list.map((r) => r.id)).toEqual(["b"]);
    expect(list[0]?.sheetId).toBe("sheet-123");
    expect(list[0]?.endCol).toBe(2);
  });

  it("manager normalization dedupes identical ranges across different ids (e.g. from concurrent inserts)", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<Y.Map<unknown>>();

    const r1 = new Y.Map<unknown>();
    r1.set("id", "a");
    r1.set("sheetId", "s1");
    r1.set("startRow", 0);
    r1.set("startCol", 0);
    r1.set("endRow", 0);
    r1.set("endCol", 0);
    r1.set("keyId", "k1");

    const r2 = new Y.Map<unknown>();
    r2.set("id", "b");
    // Identical range contents but different id.
    r2.set("sheetId", "s1");
    r2.set("startRow", 0);
    r2.set("startCol", 0);
    r2.set("endRow", 0);
    r2.set("endCol", 0);
    r2.set("keyId", "k1");

    doc.transact(() => {
      ranges.push([r1, r2]);
      metadata.set("encryptedRanges", ranges);
    });

    const mgr = new EncryptedRangeManager({ doc });
    // Both ranges are visible before normalization runs.
    expect(mgr.list().map((r) => r.id).sort()).toEqual(["a", "b"]);

    // Any tracked mutation triggers normalization (untracked) before applying the edit.
    mgr.update("a", { endCol: 1 });

    const list = mgr.list();
    expect(list.map((r) => r.id)).toEqual(["a"]);
    expect(list[0]?.endCol).toBe(1);
  });

  it("manager normalization keeps the mutated id when deduping identical ranges", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const metadata = doc.getMap("metadata");
    const ranges = new Y.Array<Y.Map<unknown>>();

    const r1 = new Y.Map<unknown>();
    r1.set("id", "a");
    r1.set("sheetId", "s1");
    r1.set("startRow", 0);
    r1.set("startCol", 0);
    r1.set("endRow", 0);
    r1.set("endCol", 0);
    r1.set("keyId", "k1");

    const r2 = new Y.Map<unknown>();
    r2.set("id", "b");
    // Identical range contents but different id.
    r2.set("sheetId", "s1");
    r2.set("startRow", 0);
    r2.set("startCol", 0);
    r2.set("endRow", 0);
    r2.set("endCol", 0);
    r2.set("keyId", "k1");

    doc.transact(() => {
      ranges.push([r1, r2]);
      metadata.set("encryptedRanges", ranges);
    });

    const mgr = new EncryptedRangeManager({ doc });

    // Mutating the second id should keep that id as the canonical survivor.
    mgr.update("b", { endCol: 2 });

    const list = mgr.list();
    expect(list.map((r) => r.id)).toEqual(["b"]);
    expect(list[0]?.endCol).toBe(2);
  });

  it("policy helper prefers the most recently added encrypted range when overlaps exist", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const mgr = new EncryptedRangeManager({ doc });
    mgr.add({ sheetId: "s1", startRow: 0, startCol: 0, endRow: 1, endCol: 1, keyId: "k1" });
    mgr.add({ sheetId: "s1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k2" });

    const policy = createEncryptionPolicyFromDoc(doc);
    expect(policy.keyIdForCell({ sheetId: "s1", row: 0, col: 0 })).toBe("k2");
    expect(policy.keyIdForCell({ sheetId: "s1", row: 1, col: 1 })).toBe("k1");
  });
});
