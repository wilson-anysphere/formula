import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("CellEditorOverlay visibility uses CSS classes (no inline style.display/zIndex)", () => {
  const filePath = path.join(__dirname, "..", "src", "editor", "cellEditorOverlay.ts");
  const content = fs.readFileSync(filePath, "utf8");

  assert.equal(
    /\.style\.display\s*=/.test(content),
    false,
    "CellEditorOverlay should not assign element.style.display; use a CSS class toggle (e.g. .cell-editor--open)",
  );
  assert.equal(
    /\.style\.zIndex\s*=/.test(content),
    false,
    "CellEditorOverlay should not assign element.style.zIndex; move stacking context into CSS",
  );

  assert.match(
    content,
    /classList\.add\(\s*["']cell-editor--open["']\s*\)/,
    "CellEditorOverlay.open should add the cell-editor--open CSS modifier class",
  );
  assert.match(
    content,
    /classList\.remove\(\s*["']cell-editor--open["']\s*\)/,
    "CellEditorOverlay.close should remove the cell-editor--open CSS modifier class",
  );
});
