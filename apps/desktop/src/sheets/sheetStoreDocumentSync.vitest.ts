import { describe, expect, test } from "vitest";

import { startSheetStoreDocumentSync } from "./sheetStoreDocumentSync";
import { WorkbookSheetStore } from "./workbookSheetStore";

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => queueMicrotask(resolve));
}

class MockDoc {
  sheetIds: string[] = [];
  private readonly listeners = new Map<string, Set<(payload: any) => void>>();

  getSheetIds(): string[] {
    return [...this.sheetIds];
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

describe("sheetStoreDocumentSync (undo/redo)", () => {
  test("syncs immediately on undo so deleted sheet ids are removed before other change listeners run", async () => {
    const doc = new MockDoc();
    doc.sheetIds = ["Sheet1", "Sheet2"];

    const store = new WorkbookSheetStore([
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
    ]);

    let activeSheetId = "Sheet2";
    const handle = startSheetStoreDocumentSync(
      doc as any,
      store,
      () => activeSheetId,
      (id) => {
        activeSheetId = id;
      },
    );

    await flushMicrotasks();
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1", "Sheet2"]);

    // Simulate undo removing the active sheet.
    doc.sheetIds = ["Sheet1"];
    doc.emit("change", { source: "undo", sheetOrderDelta: { before: ["Sheet1", "Sheet2"], after: ["Sheet1"] } });

    // For undo/redo we should sync synchronously (no microtask flush needed) so downstream
    // `document.on("change")` listeners can't recreate the deleted sheet via `getCell(...)`.
    expect(store.listAll().map((s) => s.id)).toEqual(["Sheet1"]);
    expect(activeSheetId).toBe("Sheet1");

    handle.dispose();
  });
});

