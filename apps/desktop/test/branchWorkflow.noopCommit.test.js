import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { commitIfDocumentStateChanged } from "../src/panels/branch-manager/commitIfChanged.ts";

test("commitIfDocumentStateChanged skips commit when doc state matches branch head", async () => {
  const baseline = {
    schemaVersion: 1,
    sheets: { order: ["Sheet1"], metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } } },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  /** @type {any[]} */
  const commitCalls = [];
  const branchService = {
    getCurrentState: async () => baseline,
    commit: async (...args) => {
      commitCalls.push(args);
    },
  };

  const didCommit = await commitIfDocumentStateChanged({
    actor: { userId: "u1", role: "owner" },
    branchService,
    doc: { id: "doc" },
    message: "auto: checkout",
    docToState: () => baseline,
  });

  assert.equal(didCommit, false);
  assert.equal(commitCalls.length, 0);
});

test("commitIfDocumentStateChanged commits when doc state differs from branch head", async () => {
  const baseline = {
    schemaVersion: 1,
    sheets: { order: ["Sheet1"], metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } } },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const changed = {
    ...baseline,
    cells: { Sheet1: { A1: { value: 123 } } },
  };

  /** @type {any[][]} */
  const commitCalls = [];
  const branchService = {
    getCurrentState: async () => baseline,
    commit: async (...args) => {
      commitCalls.push(args);
    },
  };

  const actor = { userId: "u1", role: "owner" };

  const didCommit = await commitIfDocumentStateChanged({
    actor,
    branchService,
    doc: { id: "doc" },
    message: "auto: checkout",
    docToState: () => changed,
  });

  assert.equal(didCommit, true);
  assert.equal(commitCalls.length, 1);
  assert.deepEqual(commitCalls[0], [actor, { nextState: changed, message: "auto: checkout" }]);
});

