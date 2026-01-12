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

