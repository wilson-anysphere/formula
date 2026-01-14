import { describe, expect, test } from "vitest";

import { startSheetStoreDocumentSync } from "./sheetStoreDocumentSync";
import { WorkbookSheetStore } from "./workbookSheetStore";

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => queueMicrotask(resolve));
}

class MockDoc {
  sheetIds: string[] = [];
  sheetMetaById: Record<string, any> = {};
  private readonly listeners = new Map<string, Set<(payload: any) => void>>();

  getSheetIds(): string[] {
    return [...this.sheetIds];
  }

  getSheetMeta(sheetId: string): any {
    return this.sheetMetaById[sheetId] ?? null;
  }

  on(event: string, listener: (payload: any) => void): () => void {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(listener);
    return () => set?.delete(listener);
  }

  emit(event: string, payload: any = {}): void {
    const set = this.listeners.get(event);
    if (!set) return;
    for (const listener of set) listener(payload);
  }
}

describe("sheetStoreDocumentSync", () => {
  test("adds missing doc sheet ids into the store (lazy sheet creation)", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1"];

    const store = new WorkbookSheetStore();
    let activeSheetId = "Sheet1";

    const handle = startSheetStoreDocumentSync(
      doc,
      store,
      () => activeSheetId,
      (id) => {
        activeSheetId = id;
      },
    );
    await flushMicrotasks();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1"]);
    expect(store.getById("Sheet1")?.name).toBe("Sheet1");

    // Add Sheet2 by mutating the doc, then emit a change.
    doc.sheetIds = ["Sheet1", "Sheet2"];
    doc.emit("change");
    await flushMicrotasks();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);
    expect(store.getById("Sheet2")?.name).toBe("Sheet2");

    handle.dispose();
  });

  test("removes store entries that no longer exist in the doc (applyState removal)", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1", "Sheet2"];

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "Sheet1";

    const handle = startSheetStoreDocumentSync(
      doc,
      store,
      () => activeSheetId,
      (id) => {
        activeSheetId = id;
      },
    );

    await flushMicrotasks();
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);

    // Simulate DocumentController.applyState ordering:
    // 1) change event fires while `getSheetIds()` still includes the soon-to-be-removed sheet
    // 2) the sheet is removed from the model after the event (same tick)
    doc.emit("change", { source: "applyState" });
    doc.sheetIds = ["Sheet1"];
    await flushMicrotasks();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1"]);

    handle.dispose();
  });

  test("removes store entries even when a last-visible-sheet guard would block (applyState removal + visibility update)", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1", "Sheet2"];
    doc.sheetMetaById = {
      Sheet1: { name: "Sheet1", visibility: "visible" },
      // DocumentController metadata may promote Sheet2 to visible during restores where it's the only
      // remaining sheet, even if the UI store still has it marked hidden.
      Sheet2: { name: "Sheet2", visibility: "visible" },
    };

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "hidden" },
    ]);

    let activeSheetId = "Sheet1";
    const handle = startSheetStoreDocumentSync(
      doc,
      store,
      () => activeSheetId,
      (id) => {
        activeSheetId = id;
      },
    );

    await flushMicrotasks();
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);

    // Simulate an applyState restore that deletes Sheet1 and makes Sheet2 visible.
    doc.emit("change", { source: "applyState" });
    doc.sheetIds = ["Sheet2"];
    doc.sheetMetaById = { Sheet2: { name: "Sheet2", visibility: "visible" } };
    await flushMicrotasks();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet2"]);
    expect(store.getById("Sheet2")?.visibility).toBe("visible");

    handle.dispose();
  });

  test("auto-activates first visible sheet if the active sheet no longer exists", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1", "Sheet2"];

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "Sheet2";
    let activated: string | null = null;

    const handle = startSheetStoreDocumentSync(
      doc,
      store,
      () => activeSheetId,
      (id) => {
        activated = id;
        activeSheetId = id;
      },
    );

    await flushMicrotasks();
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);

    // Remove the active sheet.
    doc.emit("change");
    doc.sheetIds = ["Sheet1"];
    await flushMicrotasks();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1"]);
    expect(activated).toBe("Sheet1");

    handle.dispose();
  });

  test("reorders the store to match doc sheet order when restoring via applyState", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1", "Sheet2"];

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "Sheet1";

    const handle = startSheetStoreDocumentSync(
      doc,
      store,
      () => activeSheetId,
      (id) => {
        activeSheetId = id;
      },
    );

    await flushMicrotasks();
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);

    // Simulate an applyState restore that reorders sheets.
    doc.sheetIds = ["Sheet2", "Sheet1"];
    doc.emit("change", { source: "applyState" });
    await flushMicrotasks();

    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet2", "Sheet1"]);

    handle.dispose();
  });

  test("syncs sheet metadata (name/visibility/tabColor) from the doc on applyState", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1", "Sheet2"];
    doc.sheetMetaById = {
      Sheet1: { name: "Budget", visibility: "hidden", tabColor: { rgb: "FF0000" } },
      // Keep at least one visible sheet (Excel invariant; WorkbookSheetStore enforces this too).
      Sheet2: { name: "Sheet2", visibility: "visible" },
    };

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "Sheet1";

    const handle = startSheetStoreDocumentSync(
      doc,
      store,
      () => activeSheetId,
      (id) => {
        activeSheetId = id;
      },
    );

    await flushMicrotasks();

    doc.emit("change", { source: "applyState" });
    await flushMicrotasks();

    expect(store.getById("Sheet1")?.name).toBe("Budget");
    expect(store.getById("Sheet1")?.visibility).toBe("hidden");
    expect(store.getById("Sheet1")?.tabColor?.rgb).toBe("FF0000");

    handle.dispose();
  });

  test("preserves veryHidden visibility from the doc on applyState", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1", "Sheet2"];
    doc.sheetMetaById = {
      Sheet1: { name: "Sheet1", visibility: "veryHidden" },
      Sheet2: { name: "Sheet2", visibility: "visible" },
    };

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);
    let activeSheetId = "Sheet1";

    const handle = startSheetStoreDocumentSync(
      doc,
      store,
      () => activeSheetId,
      (id) => {
        activeSheetId = id;
      },
    );

    await flushMicrotasks();

    doc.emit("change", { source: "applyState" });
    await flushMicrotasks();

    expect(store.getById("Sheet1")?.visibility).toBe("veryHidden");

    handle.dispose();
  });
});
