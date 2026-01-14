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

test("Ribbon schema includes Home → Font command ids", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const requiredIds = [
    // Font name + size presets
    "format.fontName.calibri",
    "format.fontName.arial",
    "format.fontSize.11",
    "format.fontSize.72",

    // Font size stepping
    "format.fontSize.increase",
    "format.fontSize.decrease",

    // Core toggles
    "format.toggleBold",
    "format.toggleItalic",
    "format.toggleUnderline",
    "format.toggleStrikethrough",
    "format.toggleSubscript",
    "format.toggleSuperscript",

    // Border presets (menu items)
    "format.borders.all",

    // Color presets + pickers (menu items)
    "format.fillColor.none",
    "format.fillColor.moreColors",
    "format.fontColor.automatic",
    "format.fontColor.moreColors",

    // Clear actions
    "format.clearFormats",
    "edit.clearContents",
    "format.clearAll",
  ];

  for (const id of requiredIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // Ribbon dropdown triggers (menu containers)
  const triggerIds = [
    "home.font.fontName",
    "home.font.fontSize",
    "home.font.borders",
    "home.font.fillColor",
    "home.font.fontColor",
    "home.font.clearFormatting",
  ];
  for (const id of triggerIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include trigger id ${id}`);
  }

  // Legacy Home → Font toggle ids should not appear in the schema (they were previously routed via ribbon handlers).
  const legacyIds = [
    "home.font.bold",
    "home.font.italic",
    "home.font.underline",
    "home.font.strikethrough",
    "home.font.subscript",
    "home.font.superscript",
  ];
  for (const id of legacyIds) {
    assert.doesNotMatch(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to not include legacy id ${id}`);
  }
});

test("Home → Font ribbon commands are registered in CommandRegistry and not handled via main.ts switch cases", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = fs.readFileSync(desktopCommandsPath, "utf8");

  const builtinCommandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtinCommands = fs.readFileSync(builtinCommandsPath, "utf8");

  const fontPresetsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinFormatFontCommands.ts");
  const fontPresets = fs.readFileSync(fontPresetsPath, "utf8");

  const fontDropdownPath = path.join(__dirname, "..", "src", "commands", "registerFormatFontDropdownCommands.ts");
  const fontDropdown = fs.readFileSync(fontDropdownPath, "utf8");

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = fs.readFileSync(disablingPath, "utf8");

  // Ensure registerDesktopCommands wires in the command registration modules for Home → Font.
  assert.match(desktopCommands, /\bregisterBuiltinFormatFontCommands\(/, "Expected registerDesktopCommands.ts to invoke registerBuiltinFormatFontCommands");
  assert.match(desktopCommands, /\bregisterFormatFontDropdownCommands\(/, "Expected registerDesktopCommands.ts to invoke registerFormatFontDropdownCommands");

  // Trigger ids with default actions should be registered as hidden aliases.
  const triggerAliasIds = ["home.font.borders", "home.font.fillColor", "home.font.fontColor", "home.font.fontSize"];
  for (const id of triggerAliasIds) {
    assert.match(
      desktopCommands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register hidden trigger alias ${id}`,
    );
  }

  // Commands implemented in registerDesktopCommands.ts.
  const desktopRegisteredIds = ["format.toggleStrikethrough", "format.toggleSubscript", "format.toggleSuperscript"];
  for (const id of desktopRegisteredIds) {
    assert.match(
      desktopCommands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerDesktopCommands.ts to register ${id}`,
    );
  }

  // Commands registered in the builtin command catalog.
  const builtinRegisteredIds = ["format.toggleBold", "format.toggleItalic", "format.toggleUnderline", "format.fontSize.increase", "format.fontSize.decrease"];
  for (const id of builtinRegisteredIds) {
    assert.match(
      builtinCommands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
  }

  // Font preset menu items are registered via registerBuiltinFormatFontCommands (loop over const presets).
  for (const key of ["calibri", "arial", "times", "courier"]) {
    assert.match(fontPresets, new RegExp(`\\b${escapeRegExp(key)}\\b`), `Expected registerBuiltinFormatFontCommands.ts to include font preset ${key}`);
  }
  assert.match(fontPresets, /\bformat\.fontName\./, "Expected registerBuiltinFormatFontCommands.ts to register format.fontName.* commands");
  assert.match(fontPresets, /\bformat\.fontSize\./, "Expected registerBuiltinFormatFontCommands.ts to register format.fontSize.* commands");

  // Border/color dropdown menu items are registered via registerFormatFontDropdownCommands.
  const dropdownRegisteredIds = [
    "format.borders.all",
    "format.fillColor.none",
    "format.fillColor.moreColors",
    "format.fontColor.automatic",
    "format.fontColor.moreColors",
    "format.clearFormats",
    "format.clearAll",
  ];
  for (const id of dropdownRegisteredIds) {
    assert.match(
      fontDropdown,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Expected registerFormatFontDropdownCommands.ts to reference command id ${id}`,
    );
  }

  // None of these ids should require ribbon-only exemptions.
  const exemptedIds = [...triggerAliasIds, ...desktopRegisteredIds, ...builtinRegisteredIds, ...dropdownRegisteredIds];
  for (const id of exemptedIds) {
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
});
