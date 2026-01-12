import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes the Theme selector dropdown (View → Appearance)", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

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

test("Desktop ribbon command ids for theme switching are handled in main.ts", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const expectations = [
    { commandId: "view.appearance.theme.system", preference: "system" },
    { commandId: "view.appearance.theme.light", preference: "light" },
    { commandId: "view.appearance.theme.dark", preference: "dark" },
    { commandId: "view.appearance.theme.highContrast", preference: "high-contrast" },
  ];

  for (const { commandId, preference } of expectations) {
    assert.match(
      main,
      new RegExp(
        `case\\s+["']${escapeRegExp(commandId)}["']:\\s*\\n\\s*themeController\\.setThemePreference\\(["']${escapeRegExp(preference)}["']\\)`,
        "m",
      ),
      `Expected main.ts to handle ${commandId} by calling themeController.setThemePreference("${preference}")`,
    );
  }
});

