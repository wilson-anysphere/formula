import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("main.ts wires sheetStructureHandlers into registerDesktopCommands (Insert/Delete Sheet)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const start = main.indexOf("registerDesktopCommands({");
  assert.notEqual(start, -1, "Expected main.ts to call registerDesktopCommands({ ... })");

  const end = main.indexOf("workbenchFileHandlers", start);
  assert.notEqual(end, -1, "Expected main.ts to pass workbenchFileHandlers into registerDesktopCommands");

  const segment = main.slice(start, end);

  assert.match(
    segment,
    /\bsheetStructureHandlers\s*:\s*{/,
    "Expected main.ts to pass sheetStructureHandlers into registerDesktopCommands",
  );
  assert.match(
    segment,
    /\binsertSheet\s*:\s*handleAddSheet\b/,
    "Expected sheetStructureHandlers.insertSheet to be wired to handleAddSheet",
  );
  assert.match(
    segment,
    /\bdeleteActiveSheet\s*:\s*handleDeleteActiveSheet\b/,
    "Expected sheetStructureHandlers.deleteActiveSheet to be wired to handleDeleteActiveSheet",
  );

  // These ribbon command ids are now registered in CommandRegistry; keep main.ts from
  // special-casing them in the `onUnknownCommand` switch.
  const ids = ["home.cells.insert.insertSheet", "home.cells.delete.deleteSheet"];
  for (const id of ids) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }
});

