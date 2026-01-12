import test from "node:test";
import assert from "node:assert/strict";

import { startSheetStoreDocumentSync } from "../sheetStoreDocumentSync.ts";

function flushMicrotasks() {
  return new Promise((resolve) => queueMicrotask(resolve));
}

class MockDoc {
  /** @type {string[]} */
  sheetIds = [];
  /** @type {Map<string, Set<(payload: any) => void>>} */
  listeners = new Map();

  getSheetIds() {
    return [...this.sheetIds];
  }

  on(event, listener) {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(listener);
    return () => set.delete(listener);
  }

  emit(event, payload = {}) {
    const set = this.listeners.get(event);
    if (!set) return;
    for (const listener of set) listener(payload);
  }
}

class MockSheetStore {
  /**
   * @param {{ id: string, name: string, visibility?: "visible" | "hidden" | "veryHidden" }[]} [initial]
   */
  constructor(initial = []) {
    this.sheets = initial.map((s) => ({ ...s, visibility: s.visibility ?? "visible" }));
  }

  listAll() {
    return this.sheets.map((s) => ({ ...s }));
  }

  listVisible() {
    return this.sheets.filter((s) => s.visibility === "visible").map((s) => ({ ...s }));
  }

  getById(id) {
    const sheet = this.sheets.find((s) => s.id === id);
    return sheet ? { ...sheet } : undefined;
  }

  addAfter(afterId, input) {
    const id = String(input?.id ?? "").trim();
    if (!id) throw new Error("Sheet id cannot be empty");
    const name = String(input?.name ?? id);
    const sheet = { id, name, visibility: "visible" };
    const idx = this.sheets.findIndex((s) => s.id === afterId);
    const insertIdx = idx === -1 ? this.sheets.length : idx + 1;
    this.sheets.splice(insertIdx, 0, sheet);
    return { ...sheet };
  }

  remove(id) {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) return;
    this.sheets.splice(idx, 1);
  }

  move(id, targetIndex) {
    const idx = this.sheets.findIndex((s) => s.id === id);
    if (idx === -1) throw new Error(`Sheet not found: ${id}`);
    const [sheet] = this.sheets.splice(idx, 1);
    const clamped = Math.max(0, Math.min(targetIndex, this.sheets.length));
    this.sheets.splice(clamped, 0, sheet);
  }
}

test("sheetStoreDocumentSync: adds missing doc sheet ids into the store (lazy sheet creation)", async () => {
  const doc = new MockDoc();
  doc.sheetIds = ["Sheet1"];

  const store = new MockSheetStore();
  let activeSheetId = "Sheet1";

  const handle = startSheetStoreDocumentSync(
    doc,
    /** @type {any} */ (store),
    () => activeSheetId,
    (id) => {
      activeSheetId = id;
    },
  );
  await flushMicrotasks();

  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet1"]);
  assert.equal(store.getById("Sheet1")?.name, "Sheet1");

  doc.sheetIds = ["Sheet1", "Sheet2"];
  doc.emit("change");
  await flushMicrotasks();

  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet1", "Sheet2"]);
  assert.equal(store.getById("Sheet2")?.name, "Sheet2");

  handle.dispose();
});

test("sheetStoreDocumentSync: removes store entries that no longer exist in the doc (applyState removal)", async () => {
  const doc = new MockDoc();
  doc.sheetIds = ["Sheet1", "Sheet2"];

  const store = new MockSheetStore([
    { id: "Sheet1", name: "Sheet1" },
    { id: "Sheet2", name: "Sheet2" },
  ]);

  let activeSheetId = "Sheet1";

  const handle = startSheetStoreDocumentSync(
    doc,
    /** @type {any} */ (store),
    () => activeSheetId,
    (id) => {
      activeSheetId = id;
    },
  );
  await flushMicrotasks();

  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet1", "Sheet2"]);

  // Simulate DocumentController.applyState ordering:
  // 1) change event fires while `getSheetIds()` still includes the soon-to-be-removed sheet
  // 2) the sheet is removed from the model after the event (same tick)
  doc.emit("change", { source: "applyState" });
  doc.sheetIds = ["Sheet1"];
  await flushMicrotasks();

  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet1"]);

  handle.dispose();
});

test("sheetStoreDocumentSync: auto-activates first visible sheet if the active sheet no longer exists", async () => {
  const doc = new MockDoc();
  doc.sheetIds = ["Sheet1", "Sheet2"];

  const store = new MockSheetStore([
    { id: "Sheet1", name: "Sheet1" },
    { id: "Sheet2", name: "Sheet2" },
  ]);

  let activeSheetId = "Sheet2";
  let activated = null;

  const handle = startSheetStoreDocumentSync(
    doc,
    /** @type {any} */ (store),
    () => activeSheetId,
    (id) => {
      activated = id;
      activeSheetId = id;
    },
  );

  await flushMicrotasks();
  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet1", "Sheet2"]);

  doc.emit("change");
  doc.sheetIds = ["Sheet1"];
  await flushMicrotasks();

  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet1"]);
  assert.equal(activated, "Sheet1");

  handle.dispose();
});

test("sheetStoreDocumentSync: reorders the store to match doc sheet order when restoring via applyState", async () => {
  const doc = new MockDoc();
  doc.sheetIds = ["Sheet1", "Sheet2"];

  const store = new MockSheetStore([
    { id: "Sheet1", name: "Sheet1" },
    { id: "Sheet2", name: "Sheet2" },
  ]);

  let activeSheetId = "Sheet1";

  const handle = startSheetStoreDocumentSync(
    doc,
    /** @type {any} */ (store),
    () => activeSheetId,
    (id) => {
      activeSheetId = id;
    },
  );

  await flushMicrotasks();
  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet1", "Sheet2"]);

  doc.sheetIds = ["Sheet2", "Sheet1"];
  doc.emit("change", { source: "applyState" });
  await flushMicrotasks();

  assert.deepEqual(store.listAll().map((s) => s.id), ["Sheet2", "Sheet1"]);

  handle.dispose();
});

