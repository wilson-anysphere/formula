import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("applyFormattingToSelection guards edit mode via isSpreadsheetEditing (split view aware)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const start = main.indexOf("function applyFormattingToSelection");
  assert.notEqual(start, -1, "Expected main.ts to define applyFormattingToSelection");

  const end = main.indexOf("const isReadOnly", start);
  assert.notEqual(end, -1, "Expected applyFormattingToSelection to define isReadOnly inside the function");

  const segment = main.slice(start, end);
  assert.match(
    segment,
    /\bif\s*\(\s*isSpreadsheetEditing\(\)\s*\)\s*return\s*;?/,
    "Expected applyFormattingToSelection to early-return when isSpreadsheetEditing() is true",
  );
  assert.doesNotMatch(
    segment,
    /\bif\s*\(\s*app\.isEditing\(\)\s*\)\s*return\s*;?/,
    "Did not expect applyFormattingToSelection to rely on app.isEditing() directly (should include split view secondary editor state)",
  );
});

