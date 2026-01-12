import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("SpreadsheetApp does not assign shared-grid overlay z-index via inline styles", () => {
  const spreadsheetAppPath = path.join(__dirname, "..", "src", "app", "spreadsheetApp.ts");
  const content = fs.readFileSync(spreadsheetAppPath, "utf8");

  assert.equal(
    content.includes("chartLayer.style.zIndex"),
    false,
    "SpreadsheetApp should not assign chart layer z-index via inline style"
  );
  assert.equal(
    content.includes("selectionCanvas.style.zIndex"),
    false,
    "SpreadsheetApp should not assign selection canvas z-index via inline style"
  );
});

