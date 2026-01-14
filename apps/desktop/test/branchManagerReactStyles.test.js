import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Branch manager React panels avoid inline styles", () => {
  const panelDir = path.join(__dirname, "..", "src", "panels", "branch-manager");
  const branchManagerPath = path.join(panelDir, "BranchManagerPanel.tsx");
  const mergePath = path.join(panelDir, "MergeBranchPanel.tsx");
  const cssPath = path.join(__dirname, "..", "src", "styles", "workspace.css");

  const branchManagerSource = stripComments(fs.readFileSync(branchManagerPath, "utf8"));
  const mergeSource = stripComments(fs.readFileSync(mergePath, "utf8"));
  const css = fs.readFileSync(cssPath, "utf8");

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

  // And that the CSS lives in the shared dock stylesheet (theme/token-driven).
  for (const selector of [".branch-manager", ".branch-merge"]) {
    assert.ok(css.includes(selector), `Expected workspace.css to define ${selector}`);
  }
});
