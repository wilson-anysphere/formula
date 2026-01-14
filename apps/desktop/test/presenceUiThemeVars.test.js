import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("presence collaborator pills use --formula-grid-* tokens when overlaid on the grid", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "presence-ui.css");
  const css = fs.readFileSync(cssPath, "utf8");

  assert.match(
    css,
    /\.presence-collaborators__item,[\s\S]*?\.presence-collaborators__overflow\s*\{[\s\S]*?border:\s*1px solid var\(--formula-grid-line\b/,
    "Expected collaborator pills to use --formula-grid-line for the border",
  );
  assert.match(
    css,
    /background:\s*var\(--formula-grid-header-bg\b/,
    "Expected collaborator pills to use --formula-grid-header-bg for the surface",
  );
  assert.match(
    css,
    /color:\s*var\(--formula-grid-header-text\b/,
    "Expected collaborator pills to use --formula-grid-header-text for text",
  );
});

