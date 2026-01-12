import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Branch manager React panels avoid inline styles", () => {
  const panelDir = path.join(__dirname, "..", "src", "panels", "branch-manager");
  const branchManagerPath = path.join(panelDir, "BranchManagerPanel.tsx");
  const mergePath = path.join(panelDir, "MergeBranchPanel.tsx");

  const branchManagerSource = fs.readFileSync(branchManagerPath, "utf8");
  const mergeSource = fs.readFileSync(mergePath, "utf8");

  assert.equal(
    /\bstyle\s*=/.test(branchManagerSource),
    false,
    "BranchManagerPanel.tsx should not use inline styles; styling should live in workspace.css",
  );
  assert.equal(
    /\bstyle\s*=/.test(mergeSource),
    false,
    "MergeBranchPanel.tsx should not use inline styles; styling should live in workspace.css",
  );

  // Sanity-check that the shared CSS classes are still applied.
  assert.ok(branchManagerSource.includes("branch-manager"), "BranchManagerPanel.tsx should apply the branch-manager CSS class");
  assert.ok(mergeSource.includes("branch-merge"), "MergeBranchPanel.tsx should apply the branch-merge CSS class");
});

