import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";
import { installUndoRedoShortcuts } from "../shortcuts.js";
import { installUnsavedChangesPrompt } from "../unsavedChangesGuard.js";

class FakeEventTarget {
  constructor() {
    /** @type {Map<string, Set<(event: any) => void>>} */
    this.listeners = new Map();
  }

  addEventListener(type, listener) {
    let set = this.listeners.get(type);
    if (!set) {
      set = new Set();
      this.listeners.set(type, set);
    }
    set.add(listener);
  }

  removeEventListener(type, listener) {
    this.listeners.get(type)?.delete(listener);
  }

  dispatchEvent(type, event) {
    for (const listener of this.listeners.get(type) ?? []) listener(event);
  }
}

test("Ctrl/Cmd+Z and Ctrl/Cmd+Shift+Z call undo/redo", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "x");

  const target = new FakeEventTarget();
  installUndoRedoShortcuts(target, doc);

  let prevented = 0;
  target.dispatchEvent("keydown", {
    key: "z",
    ctrlKey: true,
    shiftKey: false,
    preventDefault() {
      prevented += 1;
    },
  });
  assert.equal(doc.getCell("Sheet1", "A1").value, null);

  target.dispatchEvent("keydown", {
    key: "z",
    ctrlKey: true,
    shiftKey: true,
    preventDefault() {
      prevented += 1;
    },
  });
  assert.equal(doc.getCell("Sheet1", "A1").value, "x");
  assert.equal(prevented, 2);
});

test("beforeunload prompt triggers only when dirty", () => {
  const doc = new DocumentController();
  const target = new FakeEventTarget();
  installUnsavedChangesPrompt(target, doc, { message: "Unsaved!" });

  /** @type {{ returnValue?: any, prevented?: boolean }} */
  const cleanEvent = {};
  target.dispatchEvent("beforeunload", cleanEvent);
  assert.equal("returnValue" in cleanEvent, false);

  doc.setCellValue("Sheet1", "A1", "x");
  /** @type {{ returnValue?: any, prevented?: boolean, preventDefault?: () => void }} */
  const dirtyEvent = {
    preventDefault() {
      this.prevented = true;
    },
  };
  target.dispatchEvent("beforeunload", dirtyEvent);
  assert.equal(dirtyEvent.prevented, true);
  assert.equal(dirtyEvent.returnValue, "Unsaved!");
});

