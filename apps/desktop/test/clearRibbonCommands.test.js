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

test("Ribbon schema wires Home → Editing → Clear menu items to canonical command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const requiredIds = [
    "format.clearAll",
    "format.clearFormats",
    "edit.clearContents",
    // Unimplemented items remain ribbon-scoped for now.
    "home.editing.clear.clearComments",
    "home.editing.clear.clearHyperlinks",
  ];
  for (const id of requiredIds) {
    assert.match(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected homeTab.ts to include ${id}`,
    );
  }

  // Guardrails: do not regress to legacy ribbon-only clear ids.
  const legacyIds = [
    "home.editing.clear.clearAll",
    "home.editing.clear.clearFormats",
    "home.editing.clear.clearContents",
    // Clear Contents should route through the canonical edit command, not the legacy format variant.
    "format.clearContents",
  ];
  for (const id of legacyIds) {
    assert.doesNotMatch(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected homeTab.ts to not include legacy id ${id}`,
    );
  }

  // Ensure the canonical ids actually live under the Home → Editing → Clear dropdown (not elsewhere in the schema).
  const editingClearStart = schema.search(/\bid:\s*["']home\.editing\.clear["']/);
  assert.ok(editingClearStart >= 0, "Expected homeTab.ts to define the home.editing.clear dropdown");
  const editingClearEnd = (() => {
    const idx = schema.slice(editingClearStart).search(/\bid:\s*["']home\.editing\.sortFilter["']/);
    return idx >= 0 ? editingClearStart + idx : Math.min(schema.length, editingClearStart + 2_000);
  })();
  const editingClearBlock = schema.slice(editingClearStart, editingClearEnd);
  for (const id of requiredIds) {
    assert.match(
      editingClearBlock,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`),
      `Expected Home → Editing → Clear dropdown to include ${id}`,
    );
  }

  // Ensure the Home → Font → Clear dropdown also routes Clear Contents through the canonical edit command.
  const fontClearStart = schema.search(/\bid:\s*["']home\.font\.clearFormatting["']/);
  assert.ok(fontClearStart >= 0, "Expected homeTab.ts to define the home.font.clearFormatting dropdown");
  const fontClearEnd = (() => {
    const idx = schema.slice(fontClearStart).search(/\bid:\s*["']home\.alignment["']/);
    return idx >= 0 ? fontClearStart + idx : Math.min(schema.length, fontClearStart + 2_000);
  })();
  const fontClearBlock = schema.slice(fontClearStart, fontClearEnd);
  assert.match(fontClearBlock, /\bid:\s*["']format\.clearFormats["']/);
  assert.match(fontClearBlock, /\bid:\s*["']edit\.clearContents["']/);
  assert.match(fontClearBlock, /\bid:\s*["']format\.clearAll["']/);
});

test("Clear commands are registered under canonical ids (no legacy routing helpers)", () => {
  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = fs.readFileSync(builtinsPath, "utf8");
  const dropdownPath = path.join(__dirname, "..", "src", "commands", "registerFormatFontDropdownCommands.ts");
  const dropdown = fs.readFileSync(dropdownPath, "utf8");
  const keybindingsPath = path.join(__dirname, "..", "src", "commands", "builtinKeybindings.ts");
  const keybindings = fs.readFileSync(keybindingsPath, "utf8");
  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = fs.readFileSync(disablingPath, "utf8");
  const editingDisabledPath = path.join(__dirname, "..", "src", "ribbon", "ribbonEditingDisabledById.ts");
  const editingDisabled = fs.readFileSync(editingDisabledPath, "utf8");

  // Clear Contents is an editing command (used by Delete key + ribbon), so it should be registered as `edit.clearContents`.
  assert.match(
    builtins,
    /\bregisterBuiltinCommand\(\s*["']edit\.clearContents["']/,
    "Expected registerBuiltinCommands.ts to register edit.clearContents",
  );

  // Ribbon "Clear" menu uses canonical formatting ids for Clear Formats / Clear All.
  assert.match(
    dropdown,
    /\bregisterBuiltinCommand\(\s*["']format\.clearFormats["']/,
    "Expected registerFormatFontDropdownCommands.ts to register format.clearFormats",
  );
  assert.match(
    dropdown,
    /\bregisterBuiltinCommand\(\s*["']format\.clearAll["']/,
    "Expected registerFormatFontDropdownCommands.ts to register format.clearAll",
  );

  // format.clearContents was an older/duplicate command id and should not be reintroduced.
  assert.doesNotMatch(
    dropdown,
    /\bregisterBuiltinCommand\(\s*["']format\.clearContents["']/,
    "Expected registerFormatFontDropdownCommands.ts to not register format.clearContents",
  );

  // Delete key should dispatch through the canonical edit command.
  assert.match(
    keybindings,
    /\bcommand:\s*["']edit\.clearContents["'][\s\S]*?\bkey:\s*["']delete["']/,
    "Expected builtinKeybindings.ts to bind Delete to edit.clearContents",
  );
  assert.doesNotMatch(
    keybindings,
    /\bcommand:\s*["']format\.clearContents["']/,
    "Expected builtinKeybindings.ts to not reference legacy format.clearContents",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Clear commands should be dispatched through the CommandRegistry bridge (createRibbonActionsFromCommands),
  // not handled via main.ts's ribbon fallback switch.
  const forbiddenCases = [
    "home.editing.clear.clearAll",
    "home.editing.clear.clearFormats",
    "home.editing.clear.clearContents",
    "format.clearContents",
  ];
  for (const id of forbiddenCases) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case`,
    );
  }

  // Ensure the old bespoke routing module isn't reintroduced.
  assert.doesNotMatch(main, /\bhomeEditingClearCommandRouting\b/);
  assert.doesNotMatch(main, /\bresolveHomeEditingClearCommandTarget\b/);

  // Read-only ribbon disabling should keep Clear Contents / Clear All disabled, even for band selections.
  // `format.clearFormats` is explicitly re-enabled (formatting defaults) in read-only mode.
  const readOnlyStart = main.indexOf("const RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY");
  assert.ok(readOnlyStart >= 0, "Expected main.ts to define RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY");
  const readOnlyEnd = main.indexOf("})();", readOnlyStart);
  assert.ok(readOnlyEnd >= 0, "Expected to find end of RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY initializer");
  const readOnlyBlock = main.slice(readOnlyStart, readOnlyEnd);
  assert.match(
    readOnlyBlock,
    /delete\s+out\["format\.clearFormats"\]/,
    "Expected RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY to re-enable format.clearFormats for formatting defaults",
  );
  assert.doesNotMatch(
    readOnlyBlock,
    /delete\s+out\["edit\.clearContents"\]/,
    "Expected RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY to keep edit.clearContents disabled in read-only mode",
  );
  assert.doesNotMatch(
    readOnlyBlock,
    /delete\s+out\["format\.clearAll"\]/,
    "Expected RIBBON_DISABLED_BY_ID_WHILE_READ_ONLY to keep format.clearAll disabled in read-only mode",
  );

  // Unimplemented Clear Comments / Clear Hyperlinks should remain disabled by default.
  // Guardrail: they must not be added to the CommandRegistry exemption list.
  const exemptionsStart = disabling.indexOf("export const COMMAND_REGISTRY_EXEMPT_IDS");
  assert.ok(exemptionsStart >= 0, "Expected ribbonCommandRegistryDisabling.ts to define COMMAND_REGISTRY_EXEMPT_IDS");
  const exemptionsEnd = disabling.indexOf("]);", exemptionsStart);
  assert.ok(exemptionsEnd >= 0, "Expected to find end of COMMAND_REGISTRY_EXEMPT_IDS set");
  const exemptionsBlock = disabling.slice(exemptionsStart, exemptionsEnd);
  const notExemptIds = ["home.editing.clear.clearComments", "home.editing.clear.clearHyperlinks", "format.clearAll", "format.clearFormats", "edit.clearContents"];
  for (const id of notExemptIds) {
    assert.doesNotMatch(
      exemptionsBlock,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Expected ${id} to not be in COMMAND_REGISTRY_EXEMPT_IDS`,
    );
  }

  // Editing mode should disable Clear actions (Excel behavior).
  const shouldDisableWhileEditing = [
    "home.editing.clear",
    "format.clearAll",
    "format.clearFormats",
    "edit.clearContents",
    // Unimplemented items still show in the dropdown; ensure they're also disabled while editing.
    "home.editing.clear.clearComments",
    "home.editing.clear.clearHyperlinks",
  ];
  for (const id of shouldDisableWhileEditing) {
    assert.match(
      editingDisabled,
      new RegExp(`["']${escapeRegExp(id)}["']\\s*:\\s*true`),
      `Expected ribbonEditingDisabledById.ts to disable ${id} while editing`,
    );
  }
  assert.doesNotMatch(
    editingDisabled,
    /["']format\.clearContents["']\s*:\s*true/,
    "Expected ribbonEditingDisabledById.ts to not reference legacy format.clearContents",
  );
});
