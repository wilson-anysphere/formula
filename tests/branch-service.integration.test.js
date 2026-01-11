import test from "node:test";
import assert from "node:assert/strict";

import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { InMemoryBranchStore } from "../packages/versioning/branches/src/store/InMemoryBranchStore.js";

test("integration: create branch, diverge, merge back", async () => {
  const actor = { userId: "u1", role: "owner" };
  const store = new InMemoryBranchStore();
  const service = new BranchService({ docId: "doc1", store });

  await service.init(actor, { sheets: { Sheet1: { A1: { value: 1 } } } });

  await service.createBranch(actor, { name: "scenario" });
  await service.checkoutBranch(actor, { name: "scenario" });
  await service.commit(actor, {
    nextState: { sheets: { Sheet1: { A1: { value: 10 }, B1: { value: 99 } } } },
    message: "Scenario tweaks"
  });

  await service.checkoutBranch(actor, { name: "main" });
  await service.commit(actor, {
    nextState: { sheets: { Sheet1: { A1: { value: 5 }, C1: { value: 7 } } } },
    message: "Mainline edit"
  });

  const preview = await service.previewMerge(actor, { sourceBranch: "scenario" });
  assert.equal(preview.conflicts.length, 1, "A1 differs, should conflict");

  const merge = await service.merge(actor, {
    sourceBranch: "scenario",
    resolutions: [{ conflictIndex: 0, choice: "theirs" }]
  });

  assert.deepEqual(merge.state.cells.Sheet1, {
    A1: { value: 10 },
    B1: { value: 99 },
    C1: { value: 7 }
  });
});

test("viewer cannot commit", async () => {
  const owner = { userId: "u1", role: "owner" };
  const viewer = { userId: "u2", role: "viewer" };
  const store = new InMemoryBranchStore();
  const service = new BranchService({ docId: "doc1", store });

  await service.init(owner, { sheets: { Sheet1: { A1: { value: 1 } } } });

  await assert.rejects(
    service.commit(viewer, { nextState: { sheets: { Sheet1: { A1: { value: 2 } } } } }),
    { message: "Commit requires edit permission (role=viewer)" }
  );
});

test("commenter cannot commit", async () => {
  const owner = { userId: "u1", role: "owner" };
  const commenter = { userId: "u2", role: "commenter" };
  const store = new InMemoryBranchStore();
  const service = new BranchService({ docId: "doc1", store });

  await service.init(owner, { sheets: { Sheet1: { A1: { value: 1 } } } });

  await assert.rejects(
    service.commit(commenter, { nextState: { sheets: { Sheet1: { A1: { value: 2 } } } } }),
    { message: "Commit requires edit permission (role=commenter)" }
  );
});

test("legacy commits preserve workbook metadata (sheets/namedRanges/comments)", async () => {
  const actor = { userId: "u1", role: "owner" };
  const store = new InMemoryBranchStore();
  const service = new BranchService({ docId: "doc-legacy", store });

  await service.init(actor, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1", "Sheet2"],
      metaById: {
        Sheet1: { id: "Sheet1", name: "First" },
        Sheet2: { id: "Sheet2", name: "Second" },
      },
    },
    cells: { Sheet1: { A1: { value: 1 } }, Sheet2: {} },
    metadata: { scenario: "base" },
    namedRanges: { NR1: { sheetId: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } } },
    comments: { c1: { id: "c1", cellRef: "A1", content: "hello", resolved: false, replies: [] } },
  });

  // Old clients only know how to write the legacy `{ sheets: Record<sheetId, CellMap> }` shape.
  await service.commit(actor, { nextState: { sheets: { Sheet1: { A1: { value: 2 } } } } });

  const state = await service.getCurrentState();
  assert.equal(state.sheets.metaById.Sheet1?.name, "First");
  assert.equal(state.sheets.metaById.Sheet2?.name, "Second");
  assert.deepEqual(state.cells.Sheet1, { A1: { value: 2 } });
  assert.deepEqual(state.cells.Sheet2, {});
  assert.ok(state.sheets.order.includes("Sheet2"));
  assert.equal(state.metadata.scenario, "base");
  assert.ok(state.namedRanges.NR1);
  assert.equal(state.comments.c1?.content, "hello");
});
