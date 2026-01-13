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
  const schemaDir = path.join(__dirname, "..", "src", "ribbon", "schema");
  let schema = "";
  try {
    const tabFiles = fs
      .readdirSync(schemaDir, { withFileTypes: true })
      .filter((entry) => entry.isFile() && entry.name.endsWith(".ts"))
      .map((entry) => entry.name)
      .sort();
    schema = tabFiles.map((file) => fs.readFileSync(path.join(schemaDir, file), "utf8")).join("\n");
  } catch {
    // Back-compat: older versions kept all tab definitions in ribbonSchema.ts.
    const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
    schema = fs.readFileSync(schemaPath, "utf8");
  }

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

  // Desktop main.ts still handles the ribbon ids directly (ribbon command switch), matching
  // the Excel-style behavior used by the legacy auditing UI.
  assert.match(
    main,
    new RegExp(
      `case\\s+["']formulas\\.formulaAuditing\\.tracePrecedents["']:\\s*\\n` +
        `\\s*app\\.clearAuditing\\(\\);\\s*\\n` +
        `\\s*app\\.toggleAuditingPrecedents\\(\\);\\s*\\n` +
        `\\s*app\\.focus\\(\\);\\s*\\n` +
        `\\s*return;`,
      "m",
    ),
    "Expected main.ts to handle formulas.formulaAuditing.tracePrecedents via clearAuditing/toggleAuditingPrecedents/focus",
  );

  assert.match(
    main,
    new RegExp(
      `case\\s+["']formulas\\.formulaAuditing\\.traceDependents["']:\\s*\\n` +
        `\\s*app\\.clearAuditing\\(\\);\\s*\\n` +
        `\\s*app\\.toggleAuditingDependents\\(\\);\\s*\\n` +
        `\\s*app\\.focus\\(\\);\\s*\\n` +
        `\\s*return;`,
      "m",
    ),
    "Expected main.ts to handle formulas.formulaAuditing.traceDependents via clearAuditing/toggleAuditingDependents/focus",
  );

  assert.match(
    main,
    new RegExp(
      `case\\s+["']formulas\\.formulaAuditing\\.removeArrows["']:\\s*\\n` +
        `\\s*app\\.clearAuditing\\(\\);\\s*\\n` +
        `\\s*app\\.focus\\(\\);\\s*\\n` +
        `\\s*return;`,
      "m",
    ),
    "Expected main.ts to handle formulas.formulaAuditing.removeArrows via clearAuditing/focus",
  );

  assert.match(
    main,
    /"view\.toggleShowFormulas":\s*app\.getShowFormulas\(\)/,
    "Expected ribbon pressed state to include view.toggleShowFormulas",
  );
});
