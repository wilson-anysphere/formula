import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("FormulaBarView does not toggle visibility via element.style.display", () => {
  const filePath = path.join(__dirname, "..", "src", "formula-bar", "FormulaBarView.ts");
  const content = fs.readFileSync(filePath, "utf8");

  // Height syncing (.style.height) is allowed; visibility should be class/attribute-driven.
  assert.equal(
    content.includes(".style.display"),
    false,
    "FormulaBarView should not use element.style.display; use CSS classes or the hidden attribute instead",
  );
});

