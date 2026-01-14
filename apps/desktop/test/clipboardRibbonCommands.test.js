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
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = stripComments(fs.readFileSync(builtinsPath, "utf8"));

  // Ensure legacy ribbon-only IDs are no longer handled explicitly.
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.cut["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.copy["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.paste["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.pasteSpecial["']:/);
  assert.doesNotMatch(main, /\bcase\s+["']home\.clipboard\.pasteSpecial\.dialog["']:/);

  // Clipboard commands should be registered as built-in commands so ribbon, command palette,
  // and keybindings all share the same execution path.
  const commandIds = [
    "clipboard.cut",
    "clipboard.copy",
    "clipboard.paste",
    // Ribbon schema still includes an "All" item under Paste Special; it should be registered
    // for schema coverage but hidden from the command palette to avoid duplicate "Paste" entries.
    "clipboard.pasteSpecial.all",
    "clipboard.pasteSpecial",
    "clipboard.pasteSpecial.values",
    "clipboard.pasteSpecial.formulas",
    "clipboard.pasteSpecial.formats",
    "clipboard.pasteSpecial.transpose",
  ];
  for (const id of commandIds) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be routed via the ribbon command router)`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcommandId\\s*===\\s*["']${escapeRegExp(id)}["']`),
      `Expected main.ts to not special-case ${id} via commandId === checks (should be routed via ribbonCommandRouter)`,
    );
    assert.doesNotMatch(
      router,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected ribbonCommandRouter.ts to not handle ${id} via switch case (should dispatch via CommandRegistry)`,
    );
    assert.doesNotMatch(
      router,
      new RegExp(`\\bcommandId\\s*===\\s*["']${escapeRegExp(id)}["']`),
      `Expected ribbonCommandRouter.ts to not special-case ${id} via commandId === checks (should dispatch via CommandRegistry)`,
    );
    assert.doesNotMatch(
      router,
      new RegExp(`\\bcommandOverrides:\\s*\\{[\\s\\S]*?["']${escapeRegExp(id)}["']\\s*:`),
      `Expected ribbonCommandRouter.ts to not special-case ${id} via commandOverrides (should dispatch via CommandRegistry)`,
    );
  }

  assert.match(
    builtins,
    /\bregisterBuiltinCommand\(\s*[\s\S]*?["']clipboard\.pasteSpecial\.all["'][\s\S]*?\bwhen:\s*["']false["']/m,
    "Expected clipboard.pasteSpecial.all to be hidden via when: \"false\" (avoid duplicate command palette entries)",
  );

  // The ribbon should be mounted through the CommandRegistry bridge so registered commands
  // are executed via commandRegistry.executeCommand(...).
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
  // Guardrail: we should not reintroduce bespoke clipboard routing in the ribbon fallback.
  assert.doesNotMatch(
    router,
    /\bcommandId\.startsWith\(\s*["']clipboard\./,
    "Did not expect ribbonCommandRouter.ts to add bespoke clipboard prefix routing (dispatch should go through CommandRegistry)",
  );
  assert.doesNotMatch(
    main,
    /\bcommandId\.startsWith\(\s*["']clipboard\./,
    "Did not expect main.ts to add bespoke clipboard prefix routing (dispatch should go through ribbonCommandRouter)",
  );
});
