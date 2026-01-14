import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { readRibbonSchemaSource } from "./ribbonSchemaSource.js";
import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Insert → PivotTable command ids", () => {
  const schema = readRibbonSchemaSource("insertTab.ts");

  const ids = [
    // Canonical command (also used by View → Insert PivotTable in other UI surfaces).
    "view.insertPivotTable",
    // Ribbon-only alias (menu item: From Table/Range…).
    "insert.tables.pivotTable.fromTableRange",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected insertTab.ts to include ${id}`);
  }
});

test("Insert → PivotTable ribbon ids are registered in CommandRegistry (no exemptions / no main.ts switch cases)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = stripComments(fs.readFileSync(builtinsPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  // Canonical command should be registered.
  assert.match(
    builtins,
    /\bregisterBuiltinCommand\(\s*["']view\.insertPivotTable["']/,
    "Expected registerBuiltinCommands.ts to register view.insertPivotTable",
  );

  // Ribbon menu alias should also be registered (but hidden from the command palette).
  assert.match(
    builtins,
    /\bregisterBuiltinCommand\(\s*["']insert\.tables\.pivotTable\.fromTableRange["']/,
    "Expected registerBuiltinCommands.ts to register insert.tables.pivotTable.fromTableRange",
  );
  assert.match(
    builtins,
    /\bregisterBuiltinCommand\([\s\S]*?["']insert\.tables\.pivotTable\.fromTableRange["'][\s\S]*?\bwhen:\s*["']false["']/m,
    "Expected insert.tables.pivotTable.fromTableRange to be hidden via when: \"false\" (avoid duplicate PivotTable entries)",
  );

  const ids = ["view.insertPivotTable", "insert.tables.pivotTable.fromTableRange"];
  for (const id of ids) {
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActions)`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});

