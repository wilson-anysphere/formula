import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { readRibbonSchemaSource } from "./ribbonSchemaSource.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Home → Cells → Insert/Delete Cells command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const ids = [
    "home.cells.insert.insertCells",
    "home.cells.insert.insertSheetRows",
    "home.cells.insert.insertSheetColumns",
    "home.cells.delete.deleteCells",
    "home.cells.delete.deleteSheetRows",
    "home.cells.delete.deleteSheetColumns",
  ];
  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }
});

test("Insert/Delete Cells ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  const insertDeleteCellsIds = ["home.cells.insert.insertCells", "home.cells.delete.deleteCells"];
  for (const id of insertDeleteCellsIds) {
    assert.match(
      commands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  const sheetRowColumnIds = [
    "home.cells.insert.insertSheetRows",
    "home.cells.insert.insertSheetColumns",
    "home.cells.delete.deleteSheetRows",
    "home.cells.delete.deleteSheetColumns",
  ];
  for (const id of sheetRowColumnIds) {
    assert.match(
      commands,
      new RegExp(`\\bregisterCellsStructuralCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id} via registerCellsStructuralCommand`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActionsFromCommands\(/);
});
