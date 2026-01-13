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

test("Ribbon schema includes View/Developer macro command ids", () => {
  const schema = readRibbonSchemaSource(["viewTab.ts", "developerTab.ts"]);

  const commandIds = [
    // View → Macros.
    "view.macros.viewMacros",
    "view.macros.viewMacros.run",
    "view.macros.viewMacros.edit",
    "view.macros.viewMacros.delete",
    "view.macros.recordMacro",
    "view.macros.recordMacro.stop",
    "view.macros.useRelativeReferences",

    // Developer → Code.
    "developer.code.visualBasic",
    "developer.code.macros",
    "developer.code.macros.run",
    "developer.code.macros.edit",
    "developer.code.recordMacro",
    "developer.code.recordMacro.stop",
    "developer.code.useRelativeReferences",
    "developer.code.macroSecurity",
    "developer.code.macroSecurity.trustCenter",
  ];

  for (const commandId of commandIds) {
    assert.match(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(commandId)}["']`),
      `Expected ribbon schema to include ${commandId}`,
    );
  }
});

test("Desktop main.ts wires macro ribbon commands to Macros/Script Editor/VBA panels", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Ribbon dispatch should delegate these macro command ids through CommandRegistry so they can
  // also be used by the command palette / keybindings.
  assert.match(main, /\bisRibbonMacroCommandId\(commandId\)/);
  assert.match(main, /\bexecuteBuiltinCommand\(commandId\)/);
  assert.match(main, /\bregisterRibbonMacroCommands\(/);

  const commandIds = [
    // View → Macros.
    "view.macros.viewMacros",
    "view.macros.viewMacros.run",
    "view.macros.viewMacros.edit",
    "view.macros.viewMacros.delete",
    "view.macros.recordMacro",
    "view.macros.recordMacro.stop",
    "view.macros.useRelativeReferences",

    // Developer → Code.
    "developer.code.visualBasic",
    "developer.code.macros",
    "developer.code.macros.run",
    "developer.code.macros.edit",
    "developer.code.recordMacro",
    "developer.code.recordMacro.stop",
    "developer.code.useRelativeReferences",
    "developer.code.macroSecurity",
    "developer.code.macroSecurity.trustCenter",
  ];

  // Ensure we no longer handle these commands exclusively in the ribbon switch.
  for (const commandId of commandIds) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(commandId)}["']:`),
      `Expected main.ts to NOT handle ribbon command id ${commandId} via switch case`,
    );
  }

  // Command registration should include these ids and wire them to panels/recorder behavior.
  const registrationPath = path.join(__dirname, "..", "src", "commands", "registerRibbonMacroCommands.ts");
  const registration = fs.readFileSync(registrationPath, "utf8");

  for (const commandId of commandIds) {
    assert.match(
      registration,
      new RegExp(`["']${escapeRegExp(commandId)}["']`),
      `Expected registerRibbonMacroCommands.ts to include ${commandId}`,
    );
  }

  assert.match(registration, /\bPanelIds\.MACROS\b/, "Expected macro commands to open PanelIds.MACROS");
  assert.match(registration, /\bPanelIds\.SCRIPT_EDITOR\b/, "Expected macro Edit actions to open PanelIds.SCRIPT_EDITOR");
  assert.match(
    registration,
    /\bPanelIds\.VBA_MIGRATE\b/,
    "Expected Visual Basic to open PanelIds.VBA_MIGRATE in desktop/Tauri builds",
  );

  assert.match(registration, /\bstartMacroRecorder\(\)/, "Expected record macro commands to start the macro recorder");
  assert.match(registration, /\bstopMacroRecorder\(\)/, "Expected stop recording commands to stop the macro recorder");
});
