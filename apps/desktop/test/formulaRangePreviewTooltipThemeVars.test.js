import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("formula range preview tooltip prefers --formula-grid-* tokens", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "ui.css");
  const css = fs.readFileSync(cssPath, "utf8");

  assert.match(
    css,
    /\.formula-range-preview-tooltip\s*\{[\s\S]*?background:\s*var\(--formula-grid-bg\b/,
    "Expected tooltip surface background to use --formula-grid-bg",
  );

  assert.match(
    css,
    /\.formula-range-preview-tooltip\s*\{[\s\S]*?border:\s*1px solid var\(--formula-grid-line\b/,
    "Expected tooltip surface border to use --formula-grid-line",
  );

  assert.match(
    css,
    /\.formula-range-preview-tooltip\s*\{[\s\S]*?color:\s*var\(--formula-grid-cell-text\b/,
    "Expected tooltip surface text color to use --formula-grid-cell-text",
  );
});

