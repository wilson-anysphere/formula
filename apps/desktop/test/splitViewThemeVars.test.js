import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("split view chrome uses grid theme bridge tokens", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "workspace.css");
  const css = fs.readFileSync(cssPath, "utf8");

  assert.match(
    css,
    /\.grid-root\[data-split-active="true"\]\s*\{[\s\S]*?--formula-grid-selection-border\b/,
    "Expected active split grid ring to use --formula-grid-selection-border",
  );

  assert.match(
    css,
    /#grid-splitter\s*\{[\s\S]*?background:\s*var\(--formula-grid-line\b/,
    "Expected grid splitter to use --formula-grid-line",
  );
});

