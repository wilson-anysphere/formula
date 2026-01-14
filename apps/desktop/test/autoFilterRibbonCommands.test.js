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

test("Ribbon schema includes canonical Data → Sort & Filter AutoFilter command ids", () => {
  const homeSchema = readRibbonSchemaSource("homeTab.ts");
  const dataSchema = readRibbonSchemaSource("dataTab.ts");

  // Home tab dropdown menu items.
  const homeIds = ["data.sortFilter.filter", "data.sortFilter.clear", "data.sortFilter.reapply"];
  for (const id of homeIds) {
    assert.match(homeSchema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // Data tab group buttons + advanced menu item.
  const dataIds = [...homeIds, "data.sortFilter.advanced.clearFilter"];
  for (const id of dataIds) {
    assert.match(dataSchema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected dataTab.ts to include ${id}`);
  }
});

test("AutoFilter ribbon commands are registered in CommandRegistry (not exempted from registry-backed disabling)", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  const commandIds = [
    "data.sortFilter.filter",
    "data.sortFilter.clear",
    "data.sortFilter.reapply",
    "data.sortFilter.advanced.clearFilter",
  ];

  // Ensure the command catalog references the canonical ids.
  for (const id of commandIds) {
    assert.match(
      commands,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to reference command id ${id}`,
    );
  }
  assert.match(commands, /\bregisterBuiltinCommand\(/, "Expected registerDesktopCommands.ts to register commands");

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // main.ts should wire the AutoFilter MVP handlers into registerDesktopCommands (so the registered
  // commands can delegate to the desktop-owned filter store + prompt).
  assert.match(main, /\bregisterDesktopCommands\(/);
  assert.match(main, /\bautoFilterHandlers\s*:/);

  // The ribbon ids should be dispatched through the CommandRegistry bridge (createRibbonActionsFromCommands),
  // not handled via the `onUnknownCommand` switch in main.ts.
  for (const id of commandIds) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Since these ids are now real commands, they should not be kept in the CommandRegistry exemption list.
  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = fs.readFileSync(disablingPath, "utf8");
  for (const id of commandIds) {
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
  }
});
