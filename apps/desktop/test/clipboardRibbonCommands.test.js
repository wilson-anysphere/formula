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

test("Ribbon schema includes canonical clipboard command ids for Home → Clipboard", () => {
  const schema = readRibbonSchemaSource("homeTab.ts");

  const ids = [
    // Clipboard group core actions.
    "clipboard.cut",
    "clipboard.copy",
    "clipboard.paste",
    "clipboard.pasteSpecial",

    // Paste dropdown menu items (also used by Paste Special dropdown).
    "clipboard.pasteSpecial.values",
    "clipboard.pasteSpecial.formulas",
    "clipboard.pasteSpecial.formats",
    "clipboard.pasteSpecial.transpose",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected homeTab.ts to include ${id}`);
  }

  // Ensure the two primary controls are dropdowns.
  assert.match(schema, /\bid:\s*["']clipboard\.paste["'][\s\S]*?\bkind:\s*["']dropdown["']/);
  assert.match(schema, /\bid:\s*["']clipboard\.pasteSpecial["'][\s\S]*?\bkind:\s*["']dropdown["']/);
});

test("Desktop main.ts routes clipboard ribbon commands through the CommandRegistry", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = fs.readFileSync(builtinsPath, "utf8");

  // Ensure legacy ribbon-only IDs are no longer handled explicitly.
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.cut["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.copy["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.paste["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.pasteSpecial["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.pasteSpecial\.dialog["']:/);

  // Clipboard command IDs should be registered as built-in commands (so Ribbon + command palette + keybindings
  // share the same execution/guardrails).
  for (const id of [
    "clipboard.cut",
    "clipboard.copy",
    "clipboard.paste",
    "clipboard.pasteSpecial",
    "clipboard.pasteSpecial.values",
    "clipboard.pasteSpecial.formulas",
    "clipboard.pasteSpecial.formats",
    "clipboard.pasteSpecial.transpose",
  ]) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActionsFromCommands)`,
    );
  }
});
