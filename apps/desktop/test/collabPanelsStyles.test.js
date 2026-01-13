import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Collab Version History / Branch Manager panels are class-driven + styled via workspace.css", () => {
  const rendererPath = path.join(__dirname, "..", "src", "panels", "panelBodyRenderer.tsx");
  const versionHistoryPath = path.join(__dirname, "..", "src", "panels", "version-history", "CollabVersionHistoryPanel.tsx");
  const collabBranchManagerPath = path.join(__dirname, "..", "src", "panels", "branch-manager", "CollabBranchManagerPanel.tsx");
  const branchManagerPath = path.join(__dirname, "..", "src", "panels", "branch-manager", "BranchManagerPanel.tsx");
  const mergeBranchPath = path.join(__dirname, "..", "src", "panels", "branch-manager", "MergeBranchPanel.tsx");
  const cssPath = path.join(__dirname, "..", "src", "styles", "workspace.css");

  const renderer = fs.readFileSync(rendererPath, "utf8");
  const versionHistory = fs.readFileSync(versionHistoryPath, "utf8");
  const collabBranchManager = fs.readFileSync(collabBranchManagerPath, "utf8");
  const branchManager = fs.readFileSync(branchManagerPath, "utf8");
  const mergeBranch = fs.readFileSync(mergeBranchPath, "utf8");
  const css = fs.readFileSync(cssPath, "utf8");

  // Avoid React inline styles in the collab panels.
  for (const [fileName, source] of [
    ["panelBodyRenderer.tsx", renderer],
    ["CollabVersionHistoryPanel.tsx", versionHistory],
    ["CollabBranchManagerPanel.tsx", collabBranchManager],
    ["BranchManagerPanel.tsx", branchManager],
    ["MergeBranchPanel.tsx", mergeBranch],
  ]) {
    assert.equal(
      /\bstyle\s*=/.test(source),
      false,
      `${fileName} should not use React inline styles; collab panels should use workspace.css classes instead`,
    );
  }

  // Sanity-check that the React markup actually uses the shared classes.
  for (const className of ["collab-panel__message", "collab-panel__message--error"]) {
    assert.ok(
      renderer.includes(className) || versionHistory.includes(className) || collabBranchManager.includes(className),
      `Expected collab panels to render className="${className}"`,
    );
  }
  assert.ok(
    renderer.includes("CollabVersionHistoryPanel"),
    "Expected panelBodyRenderer.tsx to reference the CollabVersionHistoryPanel component",
  );
  assert.ok(
    renderer.includes("CollabBranchManagerPanel"),
    "Expected panelBodyRenderer.tsx to reference the CollabBranchManagerPanel component",
  );
  assert.ok(versionHistory.includes("collab-version-history"), 'Expected CollabVersionHistoryPanel.tsx to render className="collab-version-history"');
  assert.ok(branchManager.includes("branch-manager"), 'Expected BranchManagerPanel.tsx to render className="branch-manager"');
  assert.ok(mergeBranch.includes("branch-merge"), 'Expected MergeBranchPanel.tsx to render className="branch-merge"');

  const requiredSelectors = [
    // Shared message styling (loading/errors).
    ".collab-panel__message",
    ".collab-panel__message--error",
    // Version history UI.
    ".collab-version-history",
    ".collab-version-history__item",
    // Branch/merge UI.
    ".branch-manager",
    ".branch-merge",
  ];

  for (const selector of requiredSelectors) {
    assert.ok(css.includes(selector), `Expected workspace.css to define ${selector}`);
  }
});
