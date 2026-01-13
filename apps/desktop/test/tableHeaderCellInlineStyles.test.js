import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("TableHeaderCell avoids inline style blocks (use CSS classes)", () => {
  const filePath = path.join(__dirname, "..", "src", "table", "TableHeaderCell.tsx");
  const source = fs.readFileSync(filePath, "utf8");

  assert.equal(
    source.includes("style={{"),
    false,
    "TableHeaderCell should not use React inline style objects; move presentation styles to CSS",
  );

  assert.equal(
    source.includes("var(--grid-header-bg)"),
    false,
    "TableHeaderCell should not hardcode background token values; use a CSS class toggle instead",
  );

  assert.match(
    source,
    /formula-table-header-cell--styled/,
    "Expected TableHeaderCell to toggle a CSS modifier class for styled headers",
  );
});

test("TableHeaderCell header + filter button styles are defined in ui.css", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "ui.css");
  const css = fs.readFileSync(cssPath, "utf8");

  for (const selector of [
    /\.formula-table-header-cell\s*\{/,
    /\.formula-table-header-cell--styled\s*\{/,
    /\.formula-table-filter-button\s*\{/,
    /\.formula-table-filter-button:hover\s*\{/,
    /\.formula-table-filter-button:focus-visible\s*\{/,
  ]) {
    assert.ok(selector.test(css), `Expected ui.css to define ${selector}`);
  }
});

