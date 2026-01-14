import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("auditing legend overlay uses grid theme bridge tokens", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "auditing.css");
  const css = fs.readFileSync(cssPath, "utf8");

  assert.match(
    css,
    /\.auditing-legend\s*\{[\s\S]*?border:\s*1px solid var\(--formula-grid-line\b/,
    "Expected auditing legend border to use --formula-grid-line",
  );
  assert.match(
    css,
    /\.auditing-legend\s*\{[\s\S]*?background:\s*var\(--formula-grid-scrollbar-track\b/,
    "Expected auditing legend background to use --formula-grid-scrollbar-track",
  );
  assert.match(
    css,
    /\.auditing-legend\s*\{[\s\S]*?color:\s*var\(--formula-grid-cell-text\b/,
    "Expected auditing legend text to use --formula-grid-cell-text",
  );
});

