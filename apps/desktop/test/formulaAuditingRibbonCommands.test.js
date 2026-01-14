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

test("Formulas → Formula Auditing ribbon commands are registered in CommandRegistry (not wired only in main.ts)", () => {
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

  // The ribbon ids should be executed through CommandRegistry (registered built-in commands),
  // not wired directly in main.ts (ribbon command switch).
  for (const commandId of [
    "formulas.formulaAuditing.tracePrecedents",
    "formulas.formulaAuditing.traceDependents",
    "formulas.formulaAuditing.removeArrows",
  ]) {
    assert.doesNotMatch(
      main,
      new RegExp(escapeRegExp(commandId)),
      `Expected main.ts to not mention ${commandId} directly (handled via CommandRegistry)`,
    );
  }

  // Ensure the built-in command implementations call into SpreadsheetApp with Excel-like behavior.
  assert.match(
    builtins,
    new RegExp(
      `\\bregisterBuiltinCommand\\([\\s\\S]*?["']formulas\\.formulaAuditing\\.tracePrecedents["'][\\s\\S]*?\\(\\)\\s*=>\\s*\\{` +
        `[\\s\\S]*?app\\.clearAuditing\\(\\);` +
        `[\\s\\S]*?app\\.toggleAuditingPrecedents\\(\\);` +
        `[\\s\\S]*?app\\.focus\\(\\);`,
      "m",
    ),
    "Expected tracePrecedents command to clear + toggle precedents + focus SpreadsheetApp",
  );

  assert.match(
    builtins,
    new RegExp(
      `\\bregisterBuiltinCommand\\([\\s\\S]*?["']formulas\\.formulaAuditing\\.traceDependents["'][\\s\\S]*?\\(\\)\\s*=>\\s*\\{` +
        `[\\s\\S]*?app\\.clearAuditing\\(\\);` +
        `[\\s\\S]*?app\\.toggleAuditingDependents\\(\\);` +
        `[\\s\\S]*?app\\.focus\\(\\);`,
      "m",
    ),
    "Expected traceDependents command to clear + toggle dependents + focus SpreadsheetApp",
  );

  assert.match(
    builtins,
    new RegExp(
      `\\bregisterBuiltinCommand\\([\\s\\S]*?["']formulas\\.formulaAuditing\\.removeArrows["'][\\s\\S]*?\\(\\)\\s*=>\\s*\\{` +
        `[\\s\\S]*?app\\.clearAuditing\\(\\);` +
        `[\\s\\S]*?app\\.focus\\(\\);`,
      "m",
    ),
    "Expected removeArrows command to clear auditing + focus SpreadsheetApp",
  );

  assert.match(
    main,
    // Ribbon toggles are handled via createRibbonActionsFromCommands toggleOverrides.
    new RegExp(
      `toggleOverrides:\\s*\\{[\\s\\S]*?["']view\\.toggleShowFormulas["']\\s*:\\s*(?:async\\s*)?\\(pressed\\)\\s*=>\\s*\\{` +
        `[\\s\\S]*?commandRegistry\\.executeCommand\\(["']view\\.toggleShowFormulas["']`,
      "m",
    ),
    "Expected main.ts to handle view.toggleShowFormulas via the ribbon toggleOverrides hook",
  );

  assert.match(
    main,
    /"view\.toggleShowFormulas":\s*app\.getShowFormulas\(\)/,
    "Expected ribbon pressed state to include view.toggleShowFormulas",
  );

  assert.doesNotMatch(
    main,
    /formulas\.formulaAuditing\.showFormulas/,
    "Expected legacy formulas.formulaAuditing.showFormulas id to be removed from main.ts",
  );
  assert.doesNotMatch(
    main,
    /view\.show\.showFormulas/,
    "Expected legacy view.show.showFormulas id to be removed from main.ts",
  );
});
