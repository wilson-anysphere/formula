import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { SheetManager, ensureWorkbookSchema } from "../src/index.ts";

describe("@formula/collab-workbook sheet metadata", () => {
  it("ensureWorkbookSchema canonicalizes visibility + tabColor fields", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<unknown>>("sheets");

    doc.transact(() => {
      const sheet = new Y.Map<unknown>();
      sheet.set("id", "s1");
      sheet.set("name", "Sheet1");
      sheet.set("visibility", "not-a-visibility");
      sheet.set("tabColor", "ffff0000");
      sheets.push([sheet]);
    });

    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const entry = doc.getArray<Y.Map<unknown>>("sheets").get(0) as any;
    expect(entry.get("visibility")).toBe("visible");
    expect(entry.get("tabColor")).toBe("FFFF0000");
  });

  it("dedupes sheet ids deterministically and merges non-default metadata into the winner", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<unknown>>("sheets");

    // Insert a "real" Sheet1 first (hidden + tabColor + renamed), then a placeholder
    // Sheet1 later. The schema normalizer keeps the last entry by index, but should
    // not lose richer metadata from the earlier entry.
    //
    // Include another visible sheet so the schema normalizer can preserve a hidden
    // Sheet1 (workbooks must always have at least one visible sheet).
    doc.transact(() => {
      const other = new Y.Map<unknown>();
      other.set("id", "Sheet2");
      other.set("name", "Sheet2");
      other.set("visibility", "visible");
      sheets.push([other]);

      const real = new Y.Map<unknown>();
      real.set("id", "Sheet1");
      real.set("name", "Real Sheet");
      real.set("visibility", "hidden");
      real.set("tabColor", "FF00FF00");
      sheets.push([real]);

      const placeholder = new Y.Map<unknown>();
      placeholder.set("id", "Sheet1");
      placeholder.set("name", "Sheet1");
      placeholder.set("visibility", "visible");
      sheets.push([placeholder]);
    });

    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const entries = sheets.toArray() as any[];
    const sheet1Entries = entries.filter((s) => s?.get?.("id") === "Sheet1");
    expect(sheet1Entries).toHaveLength(1);
    const sheet1 = sheet1Entries[0]!;
    expect(sheet1.get("name")).toBe("Real Sheet");
    expect(sheet1.get("visibility")).toBe("hidden");
    expect(sheet1.get("tabColor")).toBe("FF00FF00");
  });

  it("preserves visibility + tabColor across rename + move operations", () => {
    const doc = new Y.Doc();
    ensureWorkbookSchema(doc, { createDefaultSheet: false });

    const mgr = new SheetManager({ doc });
    mgr.addSheet({ id: "a", name: "SheetA" });
    mgr.addSheet({ id: "b", name: "SheetB" });

    mgr.setVisibility("a", "hidden");
    mgr.setTabColor("a", "FFFF0000");
    mgr.setTabColor("b", "FF00FF00");

    mgr.renameSheet("a", "Alpha");
    expect(mgr.list().find((s) => s.id === "a")).toEqual({
      id: "a",
      name: "Alpha",
      visibility: "hidden",
      tabColor: "FFFF0000",
    });

    mgr.moveSheet("a", 1);
    const list = mgr.list();
    expect(list.map((s) => s.id)).toEqual(["b", "a"]);
    expect(list.find((s) => s.id === "a")?.visibility).toBe("hidden");
    expect(list.find((s) => s.id === "a")?.tabColor).toBe("FFFF0000");
  });
});
