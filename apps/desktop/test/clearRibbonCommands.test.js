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
});

test("Clear commands are registered under canonical ids (no legacy routing helpers)", () => {
  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = fs.readFileSync(builtinsPath, "utf8");
  const dropdownPath = path.join(__dirname, "..", "src", "commands", "registerFormatFontDropdownCommands.ts");
  const dropdown = fs.readFileSync(dropdownPath, "utf8");

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
});

