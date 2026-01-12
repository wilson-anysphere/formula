import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import {
  CollabWorkbookSheetStore,
  computeCollabSheetsKey,
  findCollabSheetIndexById,
  listSheetsFromCollabSession,
} from "../collabWorkbookSheetStore";

function makeSession(
  initial: Array<{ id: string; name?: string; visibility?: string; tabColor?: unknown }>,
): { doc: Y.Doc; sheets: Y.Array<Y.Map<unknown>>; transactLocal: (fn: () => void) => void } {
  const doc = new Y.Doc();
  const sheets = doc.getArray<Y.Map<unknown>>("sheets");
  doc.transact(() => {
    for (const sheet of initial) {
      const map = new Y.Map<unknown>();
      map.set("id", sheet.id);
      if (sheet.name !== undefined) map.set("name", sheet.name);
      if (sheet.visibility !== undefined) map.set("visibility", sheet.visibility);
      if (sheet.tabColor !== undefined) map.set("tabColor", sheet.tabColor);
      sheets.push([map]);
    }
  });
  return { doc, sheets, transactLocal: (fn) => doc.transact(fn) };
}

describe("CollabWorkbookSheetStore", () => {
  it("extracts visibility + tabColor from the collab session and canonicalizes tabColor", () => {
    const session = makeSession([
      { id: "s1", name: "Sheet1", visibility: "hidden", tabColor: "ffff0000" },
      { id: "s2", name: "Sheet2", tabColor: { rgb: "#00FF00" } },
      { id: "s3", name: "Sheet3", visibility: "veryHidden", tabColor: "#00FF00" },
    ]);

    const sheets = listSheetsFromCollabSession(session);
    expect(sheets.map((s) => [s.id, s.visibility, s.tabColor?.rgb ?? null])).toEqual([
      ["s1", "hidden", "FFFF0000"],
      ["s2", "visible", "FF00FF00"],
      ["s3", "veryHidden", "FF00FF00"],
    ]);
  });

  it("includes visibility + tabColor in the collab key so remote updates are detected", () => {
    const session = makeSession([{ id: "s1", name: "Sheet1", visibility: "visible", tabColor: "FF00FF00" }]);
    const key1 = computeCollabSheetsKey(listSheetsFromCollabSession(session));

    session.doc.transact(() => {
      const entry = session.sheets.get(0)!;
      entry.set("visibility", "hidden");
    });

    const key2 = computeCollabSheetsKey(listSheetsFromCollabSession(session));
    expect(key2).not.toBe(key1);
  });

  it("writes hide/unhide changes back to session.sheets and keeps the local store in sync", () => {
    const session = makeSession([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    const keyRef = { value: computeCollabSheetsKey(listSheetsFromCollabSession(session)) };
    const store = new CollabWorkbookSheetStore(session as any, listSheetsFromCollabSession(session), keyRef);

    store.hide("s2");
    expect(store.getById("s2")?.visibility).toBe("hidden");
    expect(findCollabSheetIndexById(session, "s2")).toBe(1);
    expect((session.sheets.get(1) as any).get("visibility")).toBe("hidden");
    expect(keyRef.value).toBe(computeCollabSheetsKey(listSheetsFromCollabSession(session)));

    store.unhide("s2");
    expect(store.getById("s2")?.visibility).toBe("visible");
    expect((session.sheets.get(1) as any).get("visibility")).toBe("visible");
    expect(keyRef.value).toBe(computeCollabSheetsKey(listSheetsFromCollabSession(session)));
  });

  it("writes setVisibility() changes back to session.sheets (including veryHidden)", () => {
    const session = makeSession([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);
    const keyRef = { value: computeCollabSheetsKey(listSheetsFromCollabSession(session)) };
    const store = new CollabWorkbookSheetStore(session as any, listSheetsFromCollabSession(session), keyRef);

    store.setVisibility("s2", "veryHidden");
    expect(store.getById("s2")?.visibility).toBe("veryHidden");
    expect((session.sheets.get(1) as any).get("visibility")).toBe("veryHidden");
    expect(keyRef.value).toBe(computeCollabSheetsKey(listSheetsFromCollabSession(session)));

    store.setVisibility("s2", "hidden");
    expect(store.getById("s2")?.visibility).toBe("hidden");
    expect((session.sheets.get(1) as any).get("visibility")).toBe("hidden");
    expect(keyRef.value).toBe(computeCollabSheetsKey(listSheetsFromCollabSession(session)));

    store.setVisibility("s2", "visible");
    expect(store.getById("s2")?.visibility).toBe("visible");
    expect((session.sheets.get(1) as any).get("visibility")).toBe("visible");
    expect(keyRef.value).toBe(computeCollabSheetsKey(listSheetsFromCollabSession(session)));
  });

  it("writes tabColor changes back to session.sheets and canonicalizes to ARGB", () => {
    const session = makeSession([{ id: "s1", name: "Sheet1", visibility: "visible" }]);
    const keyRef = { value: computeCollabSheetsKey(listSheetsFromCollabSession(session)) };
    const store = new CollabWorkbookSheetStore(session as any, listSheetsFromCollabSession(session), keyRef);

    store.setTabColor("s1", { rgb: "ffff0000" });
    expect(store.getById("s1")?.tabColor?.rgb).toBe("FFFF0000");
    expect((session.sheets.get(0) as any).get("tabColor")).toBe("FFFF0000");

    store.setTabColor("s1", undefined);
    expect(store.getById("s1")?.tabColor).toBeUndefined();
    expect((session.sheets.get(0) as any).get("tabColor")).toBeUndefined();
  });

  it("preserves visibility + tabColor metadata when renaming and moving sheets", () => {
    const session = makeSession([
      { id: "a", name: "SheetA", visibility: "hidden", tabColor: "FFFF0000" },
      { id: "b", name: "SheetB", visibility: "visible", tabColor: "FF00FF00" },
    ]);
    const keyRef = { value: computeCollabSheetsKey(listSheetsFromCollabSession(session)) };
    const store = new CollabWorkbookSheetStore(session as any, listSheetsFromCollabSession(session), keyRef);

    store.rename("a", "Alpha");
    const renamed = session.sheets.get(0) as any;
    expect(renamed.get("id")).toBe("a");
    expect(renamed.get("name")).toBe("Alpha");
    expect(renamed.get("visibility")).toBe("hidden");
    expect(renamed.get("tabColor")).toBe("FFFF0000");

    store.move("a", 1);
    expect(store.listAll().map((s) => s.id)).toEqual(["b", "a"]);

    const movedIdx = findCollabSheetIndexById(session, "a");
    expect(movedIdx).toBe(1);
    const moved = session.sheets.get(movedIdx) as any;
    expect(moved.get("visibility")).toBe("hidden");
    expect(moved.get("tabColor")).toBe("FFFF0000");

    expect(keyRef.value).toBe(computeCollabSheetsKey(listSheetsFromCollabSession(session)));
  });

  it("writes remove() changes back to session.sheets", () => {
    const session = makeSession([
      { id: "s1", name: "Sheet1", visibility: "visible" },
      { id: "s2", name: "Sheet2", visibility: "visible" },
    ]);

    const keyRef = { value: computeCollabSheetsKey(listSheetsFromCollabSession(session)) };
    const store = new CollabWorkbookSheetStore(session as any, listSheetsFromCollabSession(session), keyRef);

    store.remove("s2");
    expect(store.listAll().map((s) => s.id)).toEqual(["s1"]);
    expect(findCollabSheetIndexById(session, "s2")).toBe(-1);
    expect(session.sheets.length).toBe(1);
    expect(keyRef.value).toBe(computeCollabSheetsKey(listSheetsFromCollabSession(session)));
  });

  it("keeps veryHidden sheets out of visible UI lists while preserving them in the store", () => {
    const session = makeSession([
      { id: "v", name: "VeryHidden", visibility: "veryHidden" },
      { id: "s", name: "Shown", visibility: "visible" },
    ]);

    const store = new CollabWorkbookSheetStore(
      session as any,
      listSheetsFromCollabSession(session),
      { value: computeCollabSheetsKey(listSheetsFromCollabSession(session)) },
    );

    expect(store.listAll().map((s) => `${s.id}:${s.visibility}`)).toEqual(["v:veryHidden", "s:visible"]);
    expect(store.listVisible().map((s) => s.id)).toEqual(["s"]);
  });
});
