import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("DrawingOverlay selection chrome resolves --formula-grid-* tokens", () => {
  const filePath = path.join(__dirname, "..", "src", "drawings", "overlay.ts");
  const source = fs.readFileSync(filePath, "utf8");

  assert.match(
    source,
    /selectionStroke:\s*resolveCssVarFromStyle\(style,\s*\"--formula-grid-selection-border\"/,
    "Expected drawing overlay selection stroke to use --formula-grid-selection-border",
  );
  assert.match(
    source,
    /selectionHandleFill:\s*resolveCssVarFromStyle\(style,\s*\"--formula-grid-bg\"/,
    "Expected drawing overlay selection handle fill to use --formula-grid-bg",
  );
  assert.match(
    source,
    /placeholderLabel:\s*resolveCssVarFromStyle\(style,\s*\"--formula-grid-cell-text\"/,
    "Expected drawing overlay placeholder label color to use --formula-grid-cell-text",
  );
});

