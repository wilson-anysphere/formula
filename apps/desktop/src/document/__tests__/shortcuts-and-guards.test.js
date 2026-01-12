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

test("beforeunload prompt is suppressed when a collab session becomes active", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "x");

  /** @type {any | null} */
  let session = null;
  const app = {
    getCollabSession: () => session,
  };

  const target = new FakeEventTarget();
  installUnsavedChangesPrompt(
    target,
    {
      get isDirty() {
        if (app.getCollabSession() != null) return false;
        return doc.isDirty;
      },
    },
    { message: "Unsaved!" },
  );

  /** @type {{ returnValue?: any, prevented?: boolean, preventDefault?: () => void }} */
  const beforeCollabEvent = {
    preventDefault() {
      this.prevented = true;
    },
  };
  target.dispatchEvent("beforeunload", beforeCollabEvent);
  assert.equal(beforeCollabEvent.prevented, true);
  assert.equal(beforeCollabEvent.returnValue, "Unsaved!");

  // Simulate the collab session being created asynchronously after the prompt is installed.
  session = { id: "session" };

  /** @type {{ returnValue?: any, prevented?: boolean, preventDefault?: () => void }} */
  const afterCollabEvent = {
    preventDefault() {
      this.prevented = true;
    },
  };
  target.dispatchEvent("beforeunload", afterCollabEvent);
  assert.equal("returnValue" in afterCollabEvent, false);
  assert.equal(afterCollabEvent.prevented, undefined);
});
