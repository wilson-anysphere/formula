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
    /\.style\.display\s*=/,
    /\.style\s*\[\s*["']display["']\s*\]\s*=/,
    /\.style\.zIndex\s*=/,
    /\.style\s*\[\s*["']zIndex["']\s*\]\s*=/,
  ];

  for (const pattern of forbiddenAssignments) {
    assert.equal(
      pattern.test(content),
      false,
      `CellEditorOverlay should not assign inline styles for display/zIndex (matched ${pattern})`,
    );
  }
});
