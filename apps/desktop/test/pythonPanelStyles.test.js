import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("PythonPanel DOM mount uses CSS classes (no inline style.*)", () => {
  const filePath = path.join(__dirname, "..", "src", "panels", "python", "PythonPanel.tsx");
  const content = fs.readFileSync(filePath, "utf8");

  assert.equal(
    /\.style\./.test(content),
    false,
    "PythonPanel.tsx should not set inline styles; use CSS classes from src/styles/python-panel.css instead",
  );

  for (const needle of [
    "root.style.display",
    "toolbar.style.display",
    "output.style.background",
    "output.style.border",
    "allowlistInput.style.display",
    "allowlistInput.style.minWidth",
  ]) {
    assert.equal(content.includes(needle), false, `Expected PythonPanel.tsx to remove inline style usage: ${needle}`);
  }

  for (const className of [
    "python-panel",
    "python-panel__toolbar",
    "python-panel__network-label",
    "python-panel__allowlist-input",
    "python-panel__split",
    "python-panel__editor-wrap",
    "python-panel__editor",
    "python-panel__output",
  ]) {
    assert.ok(content.includes(className), `Expected PythonPanel.tsx to apply the ${className} CSS class`);
  }

  // Keep data-testid hooks stable for automation/e2e.
  for (const testId of [
    "python-panel-run",
    "python-clear-output",
    "python-network-permission",
    "python-network-allowlist",
    "python-panel-code",
    "python-panel-output",
  ]) {
    assert.ok(
      content.includes(`dataset.testid = "${testId}"`),
      `Expected PythonPanel.tsx to preserve data-testid="${testId}"`,
    );
  }
});

test("desktop main.ts imports python-panel.css", () => {
  const filePath = path.join(__dirname, "..", "src", "main.ts");
  const content = stripComments(fs.readFileSync(filePath, "utf8"));
  assert.match(
    content,
    /^\s*import\s+["']\.\/styles\/python-panel\.css["']\s*;?/m,
    "Expected main.ts to import ./styles/python-panel.css",
  );
});
