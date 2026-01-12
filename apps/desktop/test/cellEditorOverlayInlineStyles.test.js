import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("CellEditorOverlay avoids inline display/z-index style assignments", () => {
  const filePath = path.join(__dirname, "..", "src", "editor", "cellEditorOverlay.ts");
  const content = fs.readFileSync(filePath, "utf8");

  const forbiddenAssignments = [
    // Direct property assignments (e.g. `this.element.style.display = "none"`).
    /\.style\.display\s*=/,
    /\.style\s*\[\s*["']display["']\s*\]\s*=/,
    /\.style\.zIndex\s*=/,
    /\.style\s*\[\s*["']zIndex["']\s*\]\s*=/,
    // Alias assignments (e.g. `const style = this.element.style; style.display = "none"`).
    /\bstyle\s*\.display\s*=/,
    /\bstyle\s*\[\s*["']display["']\s*\]\s*=/,
    /\bstyle\s*\.zIndex\s*=/,
    /\bstyle\s*\[\s*["']zIndex["']\s*\]\s*=/,
    // setProperty (also mutates inline styles).
    /\.style\.setProperty\(\s*["']display["']\s*,/,
    /\bstyle\.setProperty\(\s*["']display["']\s*,/,
    /\.style\.setProperty\(\s*["']z-index["']\s*,/,
    /\bstyle\.setProperty\(\s*["']z-index["']\s*,/,
  ];

  for (const pattern of forbiddenAssignments) {
    assert.equal(
      pattern.test(content),
      false,
      `CellEditorOverlay should not assign inline styles for display/zIndex (matched ${pattern})`,
    );
  }
});
