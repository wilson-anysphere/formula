import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Ctrl/Cmd+` Show Formulas shortcut routes through CommandRegistry when available", () => {
  const filePath = path.join(__dirname, "..", "src", "app", "spreadsheetApp.ts");
  const content = fs.readFileSync(filePath, "utf8");

  // The spreadsheet keyboard handler is a legacy fallback (KeybindingService also binds Ctrl/Cmd+`),
  // but we still want it to execute the canonical command so all entry points share the same logic.
  assert.match(
    content,
    /__formulaCommandRegistry/,
    "Expected SpreadsheetApp to reference window.__formulaCommandRegistry for Show Formulas shortcut handling",
  );
  assert.match(
    content,
    /executeCommand\(["']view\.toggleShowFormulas["']\)/,
    "Expected SpreadsheetApp Show Formulas shortcut handler to execute view.toggleShowFormulas via the CommandRegistry",
  );
});

