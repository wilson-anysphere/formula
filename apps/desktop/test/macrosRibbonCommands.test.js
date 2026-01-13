import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes View/Developer macro command ids", () => {
  const viewTabPath = path.join(__dirname, "..", "src", "ribbon", "schema", "viewTab.ts");
  const developerTabPath = path.join(__dirname, "..", "src", "ribbon", "schema", "developerTab.ts");
  const schema = `${fs.readFileSync(viewTabPath, "utf8")}\n${fs.readFileSync(developerTabPath, "utf8")}`;

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

  const handledCommandIds = [
    "view.macros.viewMacros",
    "view.macros.viewMacros.run",
    "view.macros.viewMacros.edit",
    "view.macros.viewMacros.delete",
    "view.macros.recordMacro",
    "view.macros.recordMacro.stop",
    "view.macros.useRelativeReferences",
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

  for (const commandId of handledCommandIds) {
    assert.match(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(commandId)}["']:`),
      `Expected main.ts to handle ribbon command id ${commandId}`,
    );
  }

  // Panels should be opened via the shared ribbon panel helper.
  assert.match(main, /\bopenRibbonPanel\(PanelIds\.MACROS\);/, "Expected ribbon commands to open PanelIds.MACROS");
  assert.match(
    main,
    /\bopenRibbonPanel\(PanelIds\.SCRIPT_EDITOR\);/,
    "Expected macro Edit actions (best-effort) to open PanelIds.SCRIPT_EDITOR",
  );
  assert.match(
    main,
    /\bopenRibbonPanel\(PanelIds\.VBA_MIGRATE\);/,
    "Expected Visual Basic to open PanelIds.VBA_MIGRATE in desktop/Tauri builds",
  );

  // Record Macro commands should start/stop the recorder.
  assert.match(main, /\bactiveMacroRecorder\?\.\s*start\(\);/, "Expected record macro commands to call activeMacroRecorder.start()");
  assert.match(main, /\bactiveMacroRecorder\?\.\s*stop\(\);/, "Expected stop recording commands to call activeMacroRecorder.stop()");
});
