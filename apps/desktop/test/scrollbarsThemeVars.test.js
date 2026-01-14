import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("grid scrollbars are styled via --formula-grid-* tokens", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "scrollbars.css");
  const css = fs.readFileSync(cssPath, "utf8");

  assert.match(
    css,
    /\.grid-scrollbar-track\s*\{[\s\S]*?background:\s*var\(--formula-grid-scrollbar-track\b/,
    "Expected scrollbar track to use --formula-grid-scrollbar-track",
  );

  assert.match(
    css,
    /\.grid-scrollbar-thumb\s*\{[\s\S]*?background:\s*var\(--formula-grid-scrollbar-thumb\b/,
    "Expected scrollbar thumb to use --formula-grid-scrollbar-thumb",
  );

  assert.match(
    css,
    /\.grid-scrollbar-thumb:hover,[\s\S]*?--formula-grid-selection-fill\b/,
    "Expected scrollbar hover/active styles to reference --formula-grid-selection-fill",
  );

  assert.match(
    css,
    /\.grid-scrollbar-thumb:focus-visible\s*\{[\s\S]*?--formula-grid-selection-border\b/,
    "Expected scrollbar focus-visible to use --formula-grid-selection-border",
  );
});

