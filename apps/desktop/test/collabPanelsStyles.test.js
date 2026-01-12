import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Collab Version History / Branch Manager panels are class-driven + styled via workspace.css", () => {
  const rendererPath = path.join(__dirname, "..", "src", "panels", "panelBodyRenderer.tsx");
  const cssPath = path.join(__dirname, "..", "src", "styles", "workspace.css");

  const renderer = fs.readFileSync(rendererPath, "utf8");
  const css = fs.readFileSync(cssPath, "utf8");

  // Avoid React inline styles in the collab panels.
  assert.equal(
    /\bstyle\s*=/.test(renderer),
    false,
    "panelBodyRenderer.tsx should not use React inline styles; collab panels should use workspace.css classes instead",
  );

  // Sanity-check that the React markup actually uses the shared classes.
  for (const className of ["collab-panel__message", "collab-panel__message--error", "collab-version-history"]) {
    assert.ok(renderer.includes(className), `Expected panelBodyRenderer.tsx to render className=\"${className}\"`);
  }

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
