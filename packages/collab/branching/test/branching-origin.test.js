import assert from "node:assert/strict";
import test from "node:test";
import * as Y from "yjs";

import { createCollabSession } from "../../session/src/index.ts";
import { CollabBranchingWorkflow } from "../index.js";

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
