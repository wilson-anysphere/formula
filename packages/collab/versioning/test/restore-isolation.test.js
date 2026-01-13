import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabVersioning } from "../src/index.ts";

test("CollabVersioning restoreVersion does not rewind internal collaboration roots", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  // Workbook (user-visible) roots.
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");
  const metadata = doc.getMap("metadata");
  const namedRanges = doc.getMap("namedRanges");
  // A non-default but user-visible root. The adapter is designed to include
  // any instantiated roots unless excluded.
  const comments = doc.getMap("comments");

  // Internal collaboration/version-control roots that must NOT be affected by
  // snapshots/restores.
  const branches = doc.getMap("branching:branches");
  const commits = doc.getMap("branching:commits");
  const branchingMeta = doc.getMap("branching:meta");
  const cellStructuralOps = doc.getMap("cellStructuralOps");

  // Seed initial workbook state.
  metadata.set("title", "Before");
  cells.set("Sheet1:0:0", "alpha");
  namedRanges.set("MyRange", "Sheet1!A1");

  const sheet1 = new Y.Map();
  sheet1.set("id", "Sheet1");
  sheet1.set("name", "Sheet 1");
  sheets.push([sheet1]);

  const comment1 = new Y.Map();
  comment1.set("id", "c1");
  comment1.set("text", "hello");
  comments.set("c1", comment1);

  // Seed internal state.
  branches.set("main", "branch-v1");
  commits.set("c1", "commit-v1");
  branchingMeta.set("currentBranchName", "main");
  cellStructuralOps.set("op1", "before");

  const versioning = createCollabVersioning({
    // CollabVersioning only needs the session's Y.Doc for snapshot/restore.
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
    user: { userId: "user-1", userName: "User 1" },
  });
  t.after(() => versioning.destroy());

  const checkpoint = await versioning.createCheckpoint({ name: "checkpoint-1" });

  // Mutate workbook roots.
  metadata.set("title", "After");
  cells.set("Sheet1:0:0", "beta");
  namedRanges.delete("MyRange");
  comment1.set("text", "bye");

  const sheet2 = new Y.Map();
  sheet2.set("id", "Sheet2");
  sheet2.set("name", "Sheet 2");
  sheets.push([sheet2]);

  // Mutate internal roots in a way that would be visibly "rewound" if restore
  // incorrectly applied the snapshot to them.
  branches.set("main", "branch-v2");

  commits.delete("c1");
  commits.set("c2", "commit-v2");

  branchingMeta.set("currentBranchName", "feature");

  cellStructuralOps.delete("op1");
  cellStructuralOps.set("op2", "after");

  await versioning.restoreVersion(checkpoint.id);

  // Workbook state should revert to checkpoint.
  assert.equal(metadata.get("title"), "Before");
  assert.equal(cells.get("Sheet1:0:0"), "alpha");
  assert.equal(namedRanges.get("MyRange"), "Sheet1!A1");

  const restoredSheets = sheets.toArray();
  assert.equal(restoredSheets.length, 1);
  assert.equal(restoredSheets[0]?.get("id"), "Sheet1");

  assert.equal(comments.get("c1")?.get("text"), "hello");

  // Internal roots should NOT revert (they should keep post-mutation state).
  assert.equal(branches.get("main"), "branch-v2");

  assert.equal(commits.has("c1"), false);
  assert.equal(commits.get("c2"), "commit-v2");

  assert.equal(branchingMeta.get("currentBranchName"), "feature");

  assert.equal(cellStructuralOps.has("op1"), false);
  assert.equal(cellStructuralOps.get("op2"), "after");
});

test("CollabVersioning restoreVersion supports caller-provided excludeRoots (custom branch rootName)", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  // Workbook (user-visible) roots.
  const metadata = doc.getMap("metadata");
  const cells = doc.getMap("cells");

  // Internal collaboration/version-control roots with a *custom* branch root name.
  // These must be excluded via `CollabVersioningOptions.excludeRoots`.
  const branches = doc.getMap("myBranching:branches");
  const commits = doc.getMap("myBranching:commits");
  const branchingMeta = doc.getMap("myBranching:meta");

  // Seed initial workbook state.
  metadata.set("title", "Before");
  cells.set("Sheet1:0:0", "alpha");

  // Seed internal state.
  branches.set("main", "branch-v1");
  commits.set("c1", "commit-v1");
  branchingMeta.set("currentBranchName", "main");

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
    excludeRoots: ["myBranching:branches", "myBranching:commits", "myBranching:meta"],
  });
  t.after(() => versioning.destroy());

  const checkpoint = await versioning.createCheckpoint({ name: "checkpoint-1" });

  // Mutate workbook roots.
  metadata.set("title", "After");
  cells.set("Sheet1:0:0", "beta");

  // Mutate internal roots in a way that would be visibly "rewound" if restore
  // incorrectly applied the snapshot to them.
  branches.set("main", "branch-v2");
  commits.delete("c1");
  commits.set("c2", "commit-v2");
  branchingMeta.set("currentBranchName", "feature");

  await versioning.restoreVersion(checkpoint.id);

  // Workbook state should revert to checkpoint.
  assert.equal(metadata.get("title"), "Before");
  assert.equal(cells.get("Sheet1:0:0"), "alpha");

  // Custom branching roots should NOT revert (they should keep post-mutation state).
  assert.equal(branches.get("main"), "branch-v2");
  assert.equal(commits.has("c1"), false);
  assert.equal(commits.get("c2"), "commit-v2");
  assert.equal(branchingMeta.get("currentBranchName"), "feature");
});

test("CollabVersioning does not mark dirty for excluded internal root updates", async (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  // Instantiate roots before attaching VersionManager listeners so the test
  // only measures the effects of *updates*.
  const cells = doc.getMap("cells");
  const branches = doc.getMap("branching:branches");

  const versioning = createCollabVersioning({
    // @ts-expect-error - minimal session stub for unit tests
    session: { doc },
    autoStart: false,
  });
  t.after(() => versioning.destroy());

  assert.equal(versioning.manager.dirty, false);

  // Internal branch graph churn should not make the workbook dirty.
  branches.set("main", "branch-v1");
  assert.equal(versioning.manager.dirty, false);
  assert.equal(await versioning.manager.maybeSnapshot(), null);

  // Real workbook edits should still mark dirty.
  cells.set("Sheet1:0:0", "alpha");
  assert.equal(versioning.manager.dirty, true);
});
