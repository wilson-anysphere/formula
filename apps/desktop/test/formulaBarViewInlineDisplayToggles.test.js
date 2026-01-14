import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("FormulaBarView does not toggle visibility via element.style.display", () => {
  const filePath = path.join(__dirname, "..", "src", "formula-bar", "FormulaBarView.ts");
  const content = stripComments(fs.readFileSync(filePath, "utf8"));

  // Height syncing (.style.height) is allowed; visibility should be class/attribute-driven.
  for (const needle of ["textarea.style.display", "errorButton.style.display", "errorPanel.style.display"]) {
    assert.equal(
      content.includes(needle),
      false,
      `FormulaBarView should not use ${needle}; use CSS classes or the hidden attribute instead`,
    );
  }
});
