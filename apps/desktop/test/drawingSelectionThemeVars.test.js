import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("SpreadsheetApp drawing selection handles use --formula-grid-* tokens", () => {
  const filePath = path.join(__dirname, "..", "src", "app", "spreadsheetApp.ts");
  const source = fs.readFileSync(filePath, "utf8");

  // Drawings/images selection handles are painted on the grid selection canvas.
  // They should follow the grid theme bridge so theming can diverge from app chrome.
  assert.match(
    source,
    /const stroke = resolveCssVar\("--formula-grid-selection-border"/,
    "Expected drawing selection stroke to use --formula-grid-selection-border",
  );
  assert.match(
    source,
    /const handleFill = resolveCssVar\("--formula-grid-bg"/,
    "Expected drawing selection handle fill to use --formula-grid-bg",
  );

  // Guardrail: avoid regressing to app-surface tokens.
  assert.equal(
    source.includes('const stroke = resolveCssVar("--selection-border"'),
    false,
    "Drawing selection stroke should not resolve --selection-border directly (use the grid token bridge)",
  );
  assert.equal(
    source.includes('const handleFill = resolveCssVar("--bg-primary"'),
    false,
    "Drawing selection handles should not resolve --bg-primary directly (use the grid token bridge)",
  );
});

