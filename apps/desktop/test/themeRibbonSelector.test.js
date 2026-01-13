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

test("Ribbon schema includes the Theme selector dropdown (View → Appearance)", () => {
  const schema = readRibbonSchemaSource("viewTab.ts");

  // Dropdown trigger.
  assert.match(schema, /\bid:\s*["']view\.appearance\.theme["']/);
  assert.match(schema, /\bkind:\s*["']dropdown["']/);
  assert.match(schema, /\btestId:\s*["']theme-selector["']/);

  // Menu items.
  const menuItems = [
    { id: "view.appearance.theme.system", testId: "theme-option-system" },
    { id: "view.appearance.theme.light", testId: "theme-option-light" },
    { id: "view.appearance.theme.dark", testId: "theme-option-dark" },
    { id: "view.appearance.theme.highContrast", testId: "theme-option-high-contrast" },
  ];

  for (const { id, testId } of menuItems) {
    const pattern = new RegExp(
      `\\{[^}]*\\bid:\\s*["']${escapeRegExp(id)}["'][^}]*\\btestId:\\s*["']${escapeRegExp(testId)}["'][^}]*\\}`,
      "m",
    );
    assert.match(schema, pattern, `Expected ribbon schema to include ${id} with testId ${testId}`);
  }
});

test("Desktop theme switching commands are wired via registerDesktopCommands/registerBuiltinCommands", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Theme switching is wired through the shared CommandRegistry so ribbon, command palette, and
  // keybindings share the same implementation.
  assert.match(
    main,
    /\bregisterDesktopCommands\s*\(\s*\{\s*[\s\S]*?\bthemeController\b/,
    "Expected desktop startup to pass ThemeController into registerDesktopCommands",
  );
  assert.match(
    main,
    /\bregisterDesktopCommands\s*\(\s*\{\s*[\s\S]*?\brefreshRibbonUiState\s*:\s*scheduleRibbonSelectionFormatStateUpdate\b/,
    "Expected desktop startup to wire refreshRibbonUiState to scheduleRibbonSelectionFormatStateUpdate",
  );

  const desktopCommandsPath = path.join(__dirname, "..", "src", "commands", "registerDesktopCommands.ts");
  const desktopCommands = fs.readFileSync(desktopCommandsPath, "utf8");
  assert.match(
    desktopCommands,
    /\bregisterBuiltinCommands\s*\(\s*\{\s*[\s\S]*?\bthemeController\b/,
    "Expected registerDesktopCommands.ts to pass themeController into registerBuiltinCommands",
  );
  assert.match(
    desktopCommands,
    /\bregisterBuiltinCommands\s*\(\s*\{\s*[\s\S]*?\brefreshRibbonUiState\b/,
    "Expected registerDesktopCommands.ts to pass refreshRibbonUiState into registerBuiltinCommands",
  );

  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  const expectations = [
    { commandId: "view.appearance.theme.system", preference: "system" },
    { commandId: "view.appearance.theme.light", preference: "light" },
    { commandId: "view.appearance.theme.dark", preference: "dark" },
    { commandId: "view.appearance.theme.highContrast", preference: "high-contrast" },
  ];

  for (const { commandId, preference } of expectations) {
    const pattern = new RegExp(
      `commandRegistry\\.registerBuiltinCommand\\([\\s\\S]*?["']${escapeRegExp(
        commandId,
      )}["'][\\s\\S]*?themeController\\.setThemePreference\\(["']${escapeRegExp(
        preference,
      )}["']\\)[\\s\\S]*?\\brefresh\\(\\)[\\s\\S]*?\\bfocusApp\\(\\)`,
      "m",
    );
    assert.match(
      commands,
      pattern,
      `Expected registerBuiltinCommands.ts to handle ${commandId} (setThemePreference("${preference}"), refresh, focusApp)`,
    );
  }
});

test("Desktop startup instantiates and starts ThemeController in main.ts", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Import should come from the dedicated desktop theming module.
  assert.match(main, /import\s+\{\s*ThemeController\s*\}\s+from\s+["']\.\/theme\/themeController\.js["']/);

  // Instantiate and start early in startup.
  assert.match(main, /\bnew\s+ThemeController\s*\(/);
  assert.match(main, /\bthemeController\.start\s*\(\s*\)\s*;/);

  // Best-effort cleanup on unload.
  assert.match(main, /\bthemeController\.stop\s*\(\s*\)\s*;/);
});
