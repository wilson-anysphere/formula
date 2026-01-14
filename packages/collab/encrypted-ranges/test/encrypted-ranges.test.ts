import * as Y from "yjs";
import { describe, expect, it } from "vitest";
import { ensureWorkbookSchema } from "@formula/collab-workbook";

import { EncryptedRangeManager, createEncryptionPolicyFromDoc } from "../src/index.ts";

describe("@formula/collab-encrypted-ranges", () => {
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
