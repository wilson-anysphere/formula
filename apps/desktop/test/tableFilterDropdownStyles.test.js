import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("AutoFilterDropdown uses grid-scoped --formula-grid-* tokens in sort-filter.css", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "sort-filter.css");
  const css = fs.readFileSync(cssPath, "utf8");

  const expectations = [
    [
      /\.formula-table-filter-dropdown\s*\{[\s\S]*?background:\s*var\(--formula-grid-bg\b/,
      "Expected table filter dropdown surface to use --formula-grid-bg",
    ],
    [
      /\.formula-table-filter-dropdown\s*\{[\s\S]*?color:\s*var\(--formula-grid-cell-text\b/,
      "Expected table filter dropdown surface to use --formula-grid-cell-text",
    ],
    [
      /\.formula-table-filter-dropdown\s*\{[\s\S]*?border:\s*1px solid var\(--formula-grid-line\b/,
      "Expected table filter dropdown surface to use --formula-grid-line for its border",
    ],
    [
      /\.formula-table-filter-dropdown\s+\.formula-sort-filter__row:hover\s*\{[\s\S]*?--formula-grid-scrollbar-track\b/,
      "Expected table filter dropdown rows to use --formula-grid-scrollbar-track for hover",
    ],
    [
      /\.formula-table-filter-dropdown\s+\.formula-sort-filter__checkbox:focus-visible\s*\{[\s\S]*?--formula-grid-selection-border\b/,
      "Expected table filter dropdown checkboxes to use --formula-grid-selection-border for focus",
    ],
    [
      /\.formula-table-filter-dropdown\s+\.formula-sort-filter__checkbox\s*\{[\s\S]*?accent-color:\s*var\(--formula-grid-selection-border\b/,
      "Expected table filter dropdown checkboxes to use --formula-grid-selection-border for accent-color",
    ],
    [
      /\.formula-table-filter-dropdown\s+\.formula-sort-filter__button:not\(\.formula-sort-filter__button--primary\):hover:not\(:disabled\)\s*\{[\s\S]*?--formula-grid-scrollbar-track\b/,
      "Expected table filter dropdown secondary buttons to use --formula-grid-scrollbar-track for hover",
    ],
    [
      /\.formula-table-filter-dropdown\s+\.formula-sort-filter__button:not\(\.formula-sort-filter__button--primary\):active:not\(:disabled\)\s*\{[\s\S]*?--formula-grid-scrollbar-track\b/,
      "Expected table filter dropdown secondary buttons to use --formula-grid-scrollbar-track for active",
    ],
    [
      /\.formula-table-filter-dropdown\s+\.formula-sort-filter__button:focus-visible\s*\{[\s\S]*?--formula-grid-selection-fill\b/,
      "Expected table filter dropdown button focus ring to use --formula-grid-selection-fill",
    ],
  ];

  for (const [pattern, message] of expectations) {
    assert.match(css, pattern, message);
  }
});
