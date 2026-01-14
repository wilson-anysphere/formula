import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("SelectionRenderer resolves theme CSS vars relative to a provided root element", () => {
  const filePath = path.join(__dirname, "..", "src", "selection", "renderer.ts");
  const source = fs.readFileSync(filePath, "utf8");

  assert.match(
    source,
    /cssVarRoot\?:\s*HTMLElement\s*\|\s*null/,
    "Expected SelectionRendererOptions to expose cssVarRoot",
  );

  assert.match(
    source,
    /resolveCssVar\(varName,\s*\{\s*root:\s*cssVarRoot,\s*fallback\s*\}\)/,
    "Expected SelectionRenderer theme resolution to pass root: cssVarRoot into resolveCssVar",
  );
});

test("DocumentCellProvider can resolve --formula-grid-link relative to a provided cssVarRoot", () => {
  const filePath = path.join(__dirname, "..", "src", "grid", "shared", "documentCellProvider.ts");
  const source = fs.readFileSync(filePath, "utf8");

  assert.match(
    source,
    /cssVarRoot\?:\s*HTMLElement\s*\|\s*null/,
    "Expected DocumentCellProvider options to include cssVarRoot",
  );

  assert.match(
    source,
    /resolveCssVar\(\"--formula-grid-link\",\s*\{\s*root:\s*cssVarRoot,\s*fallback:\s*\"LinkText\"\s*\}\)/,
    "Expected DocumentCellProvider to pass cssVarRoot when resolving --formula-grid-link",
  );
});

test("DrawingOverlay can resolve CSS variables relative to a provided root element", () => {
  const filePath = path.join(__dirname, "..", "src", "drawings", "overlay.ts");
  const source = fs.readFileSync(filePath, "utf8");

  assert.match(
    source,
    /cssVarRoot:\s*HTMLElement\s*\|\s*null\s*=\s*null/,
    "Expected DrawingOverlay to accept a cssVarRoot constructor arg",
  );

  assert.match(
    source,
    /const root = this\.cssVarRoot \?\? /,
    "Expected DrawingOverlay to prefer cssVarRoot when reading computed styles",
  );
});

