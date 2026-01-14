import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function extractRule(css, selector) {
  // Very small utility for style-guard tests: good enough for our controlled CSS
  // (no nested braces in rule bodies).
  const re = new RegExp(`${selector}\\s*\\{([\\s\\S]*?)\\n\\}`, "m");
  return css.match(re)?.[1] ?? null;
}

test("grid root + chrome use --formula-grid-* theme tokens", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "shell.css");
  const css = fs.readFileSync(cssPath, "utf8");

  const gridRoot = extractRule(css, String.raw`\.grid-root`);
  assert.ok(gridRoot, "Expected shell.css to define a .grid-root rule");
  assert.match(gridRoot, /background:\s*var\(--formula-grid-bg\b/, "Expected grid root to use --formula-grid-bg");
  assert.match(
    gridRoot,
    /color:\s*var\(--formula-grid-cell-text\b/,
    "Expected grid root to use --formula-grid-cell-text",
  );

  const gridFocus = extractRule(css, String.raw`\.grid-root:focus-visible`);
  assert.ok(gridFocus, "Expected shell.css to define a .grid-root:focus-visible rule");
  assert.match(
    gridFocus,
    /--formula-grid-selection-border\b/,
    "Expected grid root focus ring to use --formula-grid-selection-border",
  );

  const cellEditor = extractRule(css, String.raw`\.cell-editor`);
  assert.ok(cellEditor, "Expected shell.css to define a .cell-editor rule");
  assert.match(
    cellEditor,
    /border:\s*2px solid var\(--formula-grid-selection-border\b/,
    "Expected cell editor border to use --formula-grid-selection-border",
  );
  assert.match(
    cellEditor,
    /background:\s*var\(--formula-grid-bg\b/,
    "Expected cell editor background to use --formula-grid-bg",
  );
  assert.match(
    cellEditor,
    /color:\s*var\(--formula-grid-cell-text\b/,
    "Expected cell editor text to use --formula-grid-cell-text",
  );

  const outlineToggle = extractRule(css, String.raw`\.outline-toggle`);
  assert.ok(outlineToggle, "Expected shell.css to define a .outline-toggle rule");
  assert.match(
    outlineToggle,
    /border:\s*1px solid var\(--formula-grid-line\b/,
    "Expected outline toggle border to use --formula-grid-line",
  );
  assert.match(
    outlineToggle,
    /background:\s*var\(--formula-grid-bg\b/,
    "Expected outline toggle background to use --formula-grid-bg",
  );
  assert.match(
    outlineToggle,
    /color:\s*var\(--formula-grid-cell-text\b/,
    "Expected outline toggle text to use --formula-grid-cell-text",
  );

  const outlineToggleActive = extractRule(css, String.raw`\.outline-toggle:active`);
  assert.ok(outlineToggleActive, "Expected shell.css to define a .outline-toggle:active rule");
  assert.match(
    outlineToggleActive,
    /--formula-grid-scrollbar-track\b/,
    "Expected outline toggle active state to use --formula-grid-scrollbar-track",
  );
});

