import test from "node:test";
import assert from "node:assert/strict";

import { JSDOM } from "jsdom";
import * as Y from "yjs";

import { createUndoService, REMOTE_ORIGIN } from "../packages/collab/undo/index.js";
import { FormulaConflictMonitor } from "../packages/collab/conflicts/src/formula-conflict-monitor.js";
import { ConflictUiController } from "../apps/desktop/src/collab/conflicts-ui/index.js";

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function syncDocs(docA, docB) {
  for (let i = 0; i < 10; i += 1) {
    let changed = false;
    const updateA = Y.encodeStateAsUpdate(docA, Y.encodeStateVector(docB));
    if (updateA.length > 0) {
      Y.applyUpdate(docB, updateA, REMOTE_ORIGIN);
      changed = true;
    }
    const updateB = Y.encodeStateAsUpdate(docB, Y.encodeStateVector(docA));
    if (updateB.length > 0) {
      Y.applyUpdate(docA, updateB, REMOTE_ORIGIN);
      changed = true;
    }
    if (!changed) break;
  }
}

/**
 * @param {string} userId
 * @param {HTMLElement} container
 * @param {object} [opts]
 * @param {"formula"|"formula+value"} [opts.mode]
 */
function createClient(userId, container, opts = {}) {
  const doc = new Y.Doc();
  if (typeof opts.clientID === "number") {
    doc.clientID = opts.clientID;
  }
  const cells = doc.getMap("cells");
  const origin = { type: "local", userId };
  const undo = createUndoService({ mode: "collab", doc, scope: cells, origin, captureTimeoutMs: 10_000 });

  /** @type {Array<any>} */
  const conflicts = [];

  /** @type {ConflictUiController} */
  let ui;

  const monitor = new FormulaConflictMonitor({
    doc,
    cells,
    localUserId: userId,
    origin,
    localOrigins: undo.localOrigins,
    mode: opts.mode,
    onConflict: (c) => {
      conflicts.push(c);
      ui?.addConflict(c);
    }
  });

  ui = new ConflictUiController({ container, monitor });

  return { doc, cells, monitor, conflicts, ui };
}

/**
 * @param {Y.Map<any>} cells
 * @param {string} key
 */
function getFormula(cells, key) {
  const cell = /** @type {Y.Map<any>|undefined} */ (cells.get(key));
  return (cell?.get("formula") ?? "").toString();
}

/**
 * @param {Y.Map<any>} cells
 * @param {string} key
 */
function getValue(cells, key) {
  const cell = /** @type {Y.Map<any>|undefined} */ (cells.get(key));
  return cell?.get("value") ?? null;
}

test("E2E: concurrent same-cell edit triggers conflict UI and converges after user resolution", () => {
  const dom = new JSDOM('<div id="a"></div><div id="b"></div>', { url: "http://localhost" });

  const prevWindow = globalThis.window;
  const prevDocument = globalThis.document;
  const prevEvent = globalThis.Event;

  // Expose JSDOM globals for the UI controller.
  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;

  try {
    const containerA = dom.window.document.getElementById("a");
    const containerB = dom.window.document.getElementById("b");
    assert.ok(containerA && containerB);

    const a = createClient("alice", containerA);
    const b = createClient("bob", containerB);

    // Establish base formula.
    a.monitor.setLocalFormula("s:0:0", "=1");
    syncDocs(a.doc, b.doc);

    // Offline concurrent edits.
    a.monitor.setLocalFormula("s:0:0", "=1+1");
    b.monitor.setLocalFormula("s:0:0", "=1*2");

    // Reconnect.
    syncDocs(a.doc, b.doc);

    const toastA = containerA.querySelector('[data-testid="conflict-toast"]');
    const toastB = containerB.querySelector('[data-testid="conflict-toast"]');

    assert.ok(toastA || toastB, "expected conflict toast on at least one client");

    const chosen = toastA ? a : b;
    const chosenContainer = toastA ? containerA : containerB;

    const conflict = chosen.conflicts[0];
    assert.ok(conflict, "expected recorded conflict");

    // User opens conflict dialog.
    const openBtn = chosenContainer.querySelector('[data-testid="conflict-toast-open"]');
    assert.ok(openBtn);
    openBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    const dialog = chosenContainer.querySelector('[data-testid="conflict-dialog"]');
    assert.ok(dialog, "expected conflict dialog to open");

    // User chooses "Keep yours" (local formula on the losing client).
    const keepLocalBtn = chosenContainer.querySelector('[data-testid="conflict-choose-local"]');
    assert.ok(keepLocalBtn);
    keepLocalBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    // Propagate resolution.
    syncDocs(a.doc, b.doc);

    assert.equal(getFormula(a.cells, "s:0:0"), conflict.localFormula.trim());
    assert.equal(getFormula(b.cells, "s:0:0"), conflict.localFormula.trim());

    // Toast should be gone on the resolved client.
    assert.equal(chosenContainer.querySelector('[data-testid="conflict-toast"]'), null);
  } finally {
    globalThis.window = prevWindow;
    globalThis.document = prevDocument;
    globalThis.Event = prevEvent;
  }
});

test("E2E: concurrent same-cell value edit triggers conflict UI and converges after user resolution", () => {
  const dom = new JSDOM('<div id="a"></div><div id="b"></div>', { url: "http://localhost" });

  const prevWindow = globalThis.window;
  const prevDocument = globalThis.document;
  const prevEvent = globalThis.Event;

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;

  try {
    const containerA = dom.window.document.getElementById("a");
    const containerB = dom.window.document.getElementById("b");
    assert.ok(containerA && containerB);

    const a = createClient("alice", containerA, { mode: "formula+value" });
    const b = createClient("bob", containerB, { mode: "formula+value" });

    // Establish base value.
    a.monitor.setLocalValue("s:0:0", "base");
    syncDocs(a.doc, b.doc);

    // Offline concurrent value edits.
    a.monitor.setLocalValue("s:0:0", "a");
    b.monitor.setLocalValue("s:0:0", "b");

    // Reconnect.
    syncDocs(a.doc, b.doc);

    const toastA = containerA.querySelector('[data-testid="conflict-toast"]');
    const toastB = containerB.querySelector('[data-testid="conflict-toast"]');
    assert.ok(toastA || toastB, "expected conflict toast on at least one client");

    const chosen = toastA ? a : b;
    const chosenContainer = toastA ? containerA : containerB;

    const conflict = chosen.conflicts[0];
    assert.ok(conflict, "expected recorded conflict");
    assert.equal(conflict.kind, "value");

    const openBtn = chosenContainer.querySelector('[data-testid="conflict-toast-open"]');
    assert.ok(openBtn);
    openBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    const dialog = chosenContainer.querySelector('[data-testid="conflict-dialog"]');
    assert.ok(dialog, "expected conflict dialog to open");

    const keepLocalBtn = chosenContainer.querySelector('[data-testid="conflict-choose-local"]');
    assert.ok(keepLocalBtn);
    keepLocalBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    syncDocs(a.doc, b.doc);

    assert.equal(getValue(a.cells, "s:0:0"), conflict.localValue);
    assert.equal(getValue(b.cells, "s:0:0"), conflict.localValue);

    assert.equal(chosenContainer.querySelector('[data-testid="conflict-toast"]'), null);
  } finally {
    globalThis.window = prevWindow;
    globalThis.document = prevDocument;
    globalThis.Event = prevEvent;
  }
});

test("E2E: concurrent value vs formula surfaces a content conflict and converges after user resolution", () => {
  const dom = new JSDOM('<div id="a"></div><div id="b"></div>', { url: "http://localhost" });

  const prevWindow = globalThis.window;
  const prevDocument = globalThis.document;
  const prevEvent = globalThis.Event;

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;

  try {
    const containerA = dom.window.document.getElementById("a");
    const containerB = dom.window.document.getElementById("b");
    assert.ok(containerA && containerB);

    // Deterministic tie-break: higher clientID wins map entry overwrites.
    // Ensure the formula writer wins the race (clientID 2 > 1) so the value writer
    // sees a conflict.
    const a = createClient("alice", containerA, { mode: "formula+value", clientID: 2 });
    const b = createClient("bob", containerB, { mode: "formula+value", clientID: 1 });

    // Establish a shared base cell map.
    a.monitor.setLocalValue("s:0:0", "base");
    syncDocs(a.doc, b.doc);

    // Offline concurrent edits: alice writes a formula (sets value=null), bob writes a value.
    a.monitor.setLocalFormula("s:0:0", "=1");
    b.monitor.setLocalValue("s:0:0", "bob");
    syncDocs(a.doc, b.doc);

    // Expect a content conflict on bob (the value writer).
    const toastB = containerB.querySelector('[data-testid="conflict-toast"]');
    assert.ok(toastB, "expected conflict toast on bob");

    const conflict = b.conflicts[0];
    assert.ok(conflict, "expected recorded conflict");
    assert.equal(conflict.kind, "content");
    assert.equal(conflict.local.type, "value");
    assert.equal(conflict.remote.type, "formula");

    // Concurrent formula should be present at conflict time.
    assert.equal(getFormula(b.cells, "s:0:0"), "=1");

    const openBtn = containerB.querySelector('[data-testid="conflict-toast-open"]');
    assert.ok(openBtn);
    openBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    const dialog = containerB.querySelector('[data-testid="conflict-dialog"]');
    assert.ok(dialog, "expected conflict dialog to open");

    // Choose the local value (keep yours) - should clear the formula.
    const keepOursBtn = containerB.querySelector('[data-testid="conflict-choose-local"]');
    assert.ok(keepOursBtn);
    keepOursBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    syncDocs(a.doc, b.doc);

    assert.equal(getFormula(a.cells, "s:0:0"), "");
    assert.equal(getFormula(b.cells, "s:0:0"), "");
    assert.equal(getValue(a.cells, "s:0:0"), "bob");
    assert.equal(getValue(b.cells, "s:0:0"), "bob");

    // Toast should be gone on the resolved client.
    assert.equal(containerB.querySelector('[data-testid="conflict-toast"]'), null);
  } finally {
    globalThis.window = prevWindow;
    globalThis.document = prevDocument;
    globalThis.Event = prevEvent;
  }
});

test("E2E: concurrent value vs formula (value wins) surfaces a content conflict and converges after user resolution", () => {
  const dom = new JSDOM('<div id="a"></div><div id="b"></div>', { url: "http://localhost" });

  const prevWindow = globalThis.window;
  const prevDocument = globalThis.document;
  const prevEvent = globalThis.Event;

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Event = dom.window.Event;

  try {
    const containerA = dom.window.document.getElementById("a");
    const containerB = dom.window.document.getElementById("b");
    assert.ok(containerA && containerB);

    // Deterministic tie-break: value writer has higher clientID and wins.
    const a = createClient("alice", containerA, { mode: "formula+value", clientID: 1 });
    const b = createClient("bob", containerB, { mode: "formula+value", clientID: 2 });

    a.monitor.setLocalValue("s:0:0", "base");
    syncDocs(a.doc, b.doc);

    // Offline concurrent edits: alice writes a formula; bob writes a value.
    a.monitor.setLocalFormula("s:0:0", "=1");
    b.monitor.setLocalValue("s:0:0", "bob");
    syncDocs(a.doc, b.doc);

    // Expect conflict toast on alice (the formula writer).
    const toastA = containerA.querySelector('[data-testid="conflict-toast"]');
    assert.ok(toastA, "expected conflict toast on alice");

    const conflict = a.conflicts[0];
    assert.ok(conflict, "expected recorded conflict");
    assert.equal(conflict.kind, "content");
    assert.equal(conflict.local.type, "formula");
    assert.equal(conflict.remote.type, "value");

    const openBtn = containerA.querySelector('[data-testid="conflict-toast-open"]');
    assert.ok(openBtn);
    openBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    const dialog = containerA.querySelector('[data-testid="conflict-dialog"]');
    assert.ok(dialog, "expected conflict dialog to open");

    // Choose the remote value (already applied) - should converge without clobbering.
    const useTheirsBtn = containerA.querySelector('[data-testid="conflict-choose-remote"]');
    assert.ok(useTheirsBtn);
    useTheirsBtn.dispatchEvent(new dom.window.Event("click", { bubbles: true }));

    syncDocs(a.doc, b.doc);

    assert.equal(getFormula(a.cells, "s:0:0"), "");
    assert.equal(getFormula(b.cells, "s:0:0"), "");
    assert.equal(getValue(a.cells, "s:0:0"), "bob");
    assert.equal(getValue(b.cells, "s:0:0"), "bob");

    assert.equal(containerA.querySelector('[data-testid="conflict-toast"]'), null);
  } finally {
    globalThis.window = prevWindow;
    globalThis.document = prevDocument;
    globalThis.Event = prevEvent;
  }
});
