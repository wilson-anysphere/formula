import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Formulas → Formula Auditing command ids", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

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

  assert.match(
    main,
    new RegExp(
      `case\\s+["']formulas\\.formulaAuditing\\.tracePrecedents["']:\\s*\\n` +
        `\\s*app\\.clearAuditing\\(\\);\\s*\\n` +
        `\\s*app\\.toggleAuditingPrecedents\\(\\);\\s*\\n` +
        `\\s*app\\.focus\\(\\);`,
      "m",
    ),
    'Expected main.ts to handle formulas.formulaAuditing.tracePrecedents via clearAuditing/toggleAuditingPrecedents/focus',
  );

  assert.match(
    main,
    new RegExp(
      `case\\s+["']formulas\\.formulaAuditing\\.traceDependents["']:\\s*\\n` +
        `\\s*app\\.clearAuditing\\(\\);\\s*\\n` +
        `\\s*app\\.toggleAuditingDependents\\(\\);\\s*\\n` +
        `\\s*app\\.focus\\(\\);`,
      "m",
    ),
    'Expected main.ts to handle formulas.formulaAuditing.traceDependents via clearAuditing/toggleAuditingDependents/focus',
  );

  assert.match(
    main,
    new RegExp(
      `case\\s+["']formulas\\.formulaAuditing\\.removeArrows["']:\\s*\\n` +
        `\\s*app\\.clearAuditing\\(\\);\\s*\\n` +
        `\\s*app\\.focus\\(\\);`,
      "m",
    ),
    'Expected main.ts to handle formulas.formulaAuditing.removeArrows via clearAuditing/focus',
  );

  assert.match(
    main,
    new RegExp(
      `case\\s+["']view\\.toggleShowFormulas["']:\\s*(?:\\{\\s*)?` +
        // Ensure the ribbon toggle routes through the canonical command handler.
        `[\\s\\S]*?commandRegistry\\.executeCommand\\(["']view\\.toggleShowFormulas["']\\)`,
      "m",
    ),
    "Expected main.ts to handle view.toggleShowFormulas via commandRegistry.executeCommand(view.toggleShowFormulas)",
  );

  assert.match(
    main,
    /"view\.toggleShowFormulas":\s*app\.getShowFormulas\(\)/,
    "Expected ribbon pressed state to include view.toggleShowFormulas",
  );

  assert.match(
    main,
    /commandId\s*===\s*["']view\.toggleShowFormulas["']/,
    "Expected onCommand to ignore view.toggleShowFormulas (toggle semantics handled in onToggle)",
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
