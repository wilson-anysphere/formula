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

test("Ribbon schema includes Formulas → Formula Auditing command ids", () => {
  const schema = readRibbonSchemaSource();

  const commandIds = [
    "formulas.formulaAuditing.tracePrecedents",
    "formulas.formulaAuditing.traceDependents",
    "formulas.formulaAuditing.removeArrows",
    // "Show Formulas" is intentionally unified with the canonical view command.
    "view.toggleShowFormulas",
  ];

  for (const commandId of commandIds) {
    assert.match(
      schema,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(commandId)}["']`),
      `Expected ribbon schema to include ${commandId}`,
    );
  }
});

test("Desktop main.ts wires Formulas → Formula Auditing commands to SpreadsheetApp", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = fs.readFileSync(builtinsPath, "utf8");

  // Coverage: the ribbon schema uses these ids; ensure we register builtins using the same ids
  // (so the ribbon + command palette share the same wiring).
  for (const commandId of [
    "formulas.formulaAuditing.tracePrecedents",
    "formulas.formulaAuditing.traceDependents",
    "formulas.formulaAuditing.removeArrows",
  ]) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*\\n\\s*["']${escapeRegExp(commandId)}["']`),
      `Expected registerBuiltinCommands.ts to register ${commandId}`,
    );
  }

  // The desktop shell now routes ribbon commands through CommandRegistry when a builtin command exists
  // with the same id. These ribbon command ids should *not* be special-cased in main.ts anymore.
  for (const commandId of [
    "formulas.formulaAuditing.tracePrecedents",
    "formulas.formulaAuditing.traceDependents",
    "formulas.formulaAuditing.removeArrows",
  ]) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\b${escapeRegExp(commandId)}\\b`),
      `Did not expect main.ts to special-case ${commandId}; it should be executed via CommandRegistry`,
    );
  }

  assert.match(
    main,
    /const\s+cmd\s*=\s*commandRegistry\.getCommand\(commandId\)[\s\S]*?cmd\?\.\s*source\.kind\s*===\s*["']builtin["'][\s\S]*?executeBuiltinCommand\(commandId\);/m,
    "Expected main.ts ribbon handler to dispatch builtin ribbon ids via CommandRegistry",
  );

  assert.match(
    main,
    /"view\.toggleShowFormulas":\s*app\.getShowFormulas\(\)/,
    "Expected ribbon pressed state to include view.toggleShowFormulas",
  );
});
