import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";
import * as Y from "yjs";

import { createCollabSession } from "../../session/src/index.ts";
import { CollabBranchingWorkflow } from "../index.js";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

/**
 * @param {string} value
 * @returns {import("../../../versioning/branches/src/types.js").DocumentState}
 */
function makeState(value) {
  return {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } },
    },
    cells: {
      Sheet1: {
        A1: { value },
      },
    },
    metadata: {},
    namedRanges: {},
    comments: {},
  };
}

test("CollabBranchingWorkflow: checkout/merge default origin is not undo-tracked", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, undo: {} });
  assert.ok(session.undo, "expected undo to be enabled");
  assert.equal(session.undo.canUndo(), false);

  /** @type {any[]} */
  const origins = [];
  doc.on("update", (_update, origin) => origins.push(origin));

  /** @type {any} */
  const branchService = {
    async checkoutBranch(_actor, _input) {
      return makeState("checkout");
    },
    async merge(_actor, _input) {
      return { state: makeState("merge") };
    },
  };

  const workflow = new CollabBranchingWorkflow({ session, branchService });
  const actor = { userId: "u1", role: "owner" };
  await workflow.checkoutBranch(actor, { name: "feature" });
  await workflow.merge(actor, { sourceBranch: "feature", resolutions: [] });

  assert.ok(origins.includes("branching-apply"), "expected checkout/merge to apply with origin \"branching-apply\"");
  assert.equal(origins.includes(session.origin), false, "expected checkout/merge not to apply with session.origin by default");
  assert.equal(session.undo.canUndo(), false);
});

test("CollabBranchingWorkflow: checkout/merge can opt into session.origin for undo", async () => {
  const doc = new Y.Doc();
  const session = createCollabSession({ doc, undo: {} });
  assert.ok(session.undo, "expected undo to be enabled");
  assert.equal(session.undo.canUndo(), false);

  /** @type {any[]} */
  const origins = [];
  doc.on("update", (_update, origin) => origins.push(origin));

  /** @type {any} */
  const branchService = {
    async checkoutBranch(_actor, _input) {
      return makeState("checkout");
    },
    async merge(_actor, _input) {
      return { state: makeState("merge") };
    },
  };

  const workflow = new CollabBranchingWorkflow({ session, branchService, applyWithSessionOrigin: true });
  const actor = { userId: "u1", role: "owner" };
  await workflow.checkoutBranch(actor, { name: "feature" });
  await workflow.merge(actor, { sourceBranch: "feature", resolutions: [] });

  assert.ok(origins.includes(session.origin), "expected checkout/merge to apply with session.origin when opted in");
  assert.equal(session.undo.canUndo(), true);
});

test("CollabBranchingWorkflow.getCurrentBranchName works when branching roots were created by a different Yjs instance (CJS getMap)", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();
  const session = createCollabSession({ doc });

  // Simulate a mixed module loader environment where another Yjs instance eagerly
  // instantiates the branching roots before this workflow code touches them.
  const meta = Ycjs.Doc.prototype.getMap.call(doc, "branching:meta");
  const branches = Ycjs.Doc.prototype.getMap.call(doc, "branching:branches");
  meta.set("currentBranchName", "feature");
  branches.set("feature", 1);

  assert.throws(() => doc.getMap("branching:meta"), /different constructor/);
  assert.throws(() => doc.getMap("branching:branches"), /different constructor/);

  /** @type {any} */
  const branchService = {};
  const workflow = new CollabBranchingWorkflow({ session, branchService });

  assert.equal(workflow.getCurrentBranchName(), "feature");

  session.destroy();
  doc.destroy();
});

test("CollabBranchingWorkflow.getCurrentBranchName works when branching roots were created by a different Yjs instance (CJS Doc.get placeholder)", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();
  const session = createCollabSession({ doc });

  // Simulate another Yjs module instance touching the roots via `Doc.get(name)`
  // (defaulting to `AbstractType`), leaving foreign placeholder constructors.
  Ycjs.Doc.prototype.get.call(doc, "branching:meta");
  Ycjs.Doc.prototype.get.call(doc, "branching:branches");

  // Hydrate content via a CJS update so the placeholders have data.
  const remote = new Ycjs.Doc();
  remote.getMap("branching:meta").set("currentBranchName", "feature");
  remote.getMap("branching:branches").set("feature", 1);
  const update = Ycjs.encodeStateAsUpdate(remote);
  Ycjs.applyUpdate(doc, update);

  // Regression: `doc.getMap(...)` from the ESM build throws "different constructor"
  // when the root placeholder was created by a different Yjs module instance.
  assert.throws(() => doc.getMap("branching:meta"), /different constructor/);
  assert.throws(() => doc.getMap("branching:branches"), /different constructor/);

  /** @type {any} */
  const branchService = {};
  const workflow = new CollabBranchingWorkflow({ session, branchService });

  assert.equal(workflow.getCurrentBranchName(), "feature");
  assert.ok(doc.getMap("branching:meta") instanceof Y.Map);
  assert.ok(doc.getMap("branching:branches") instanceof Y.Map);

  session.destroy();
  doc.destroy();
  remote.destroy();
});
