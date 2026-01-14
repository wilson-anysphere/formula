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

test("Ribbon schema includes Data → Forecast → What-If Analysis command ids", () => {
  const schema = readRibbonSchemaSource("dataTab.ts");

  // Dropdown trigger id (UI-only).
  assert.match(schema, /\bid:\s*["']data\.forecast\.whatIfAnalysis["']/);
  assert.match(schema, /\btestId:\s*["']ribbon-what-if-analysis["']/);

  // Menu item ids (real commands).
  const ids = [
    "data.forecast.whatIfAnalysis.scenarioManager",
    "data.forecast.whatIfAnalysis.goalSeek",
    "data.forecast.whatIfAnalysis.monteCarlo",
    // Intentionally unimplemented (disabled by default via CommandRegistry registration).
    "data.forecast.whatIfAnalysis.dataTable",
  ];

  for (const id of ids) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected dataTab.ts to include ${id}`);
  }
});

test("What-If Analysis ribbon commands are registered in CommandRegistry (and avoid ribbon-only exemptions)", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));

  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = stripComments(fs.readFileSync(builtinsPath, "utf8"));

  const disablingPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRegistryDisabling.ts");
  const disabling = stripComments(fs.readFileSync(disablingPath, "utf8"));

  // Desktop wiring: Goal Seek depends on main.ts passing the dialog opener into registerDesktopCommands.
  assert.match(
    main,
    /\bregisterDesktopCommands\s*\(\s*\{[\s\S]*?\bopenGoalSeekDialog\b/,
    "Expected desktop startup to pass openGoalSeekDialog into registerDesktopCommands",
  );

  const implementedIds = [
    "data.forecast.whatIfAnalysis.scenarioManager",
    "data.forecast.whatIfAnalysis.monteCarlo",
    "data.forecast.whatIfAnalysis.goalSeek",
  ];

  for (const id of implementedIds) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
    assert.doesNotMatch(
      disabling,
      new RegExp(`["']${escapeRegExp(id)}["']`),
      `Did not expect ribbonCommandRegistryDisabling.ts to exempt implemented command id ${id}`,
    );
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(id)}["']:`),
      `Expected main.ts to not handle ${id} via switch case (should be dispatched by createRibbonActions)`,
    );
  }

  // Data Table is intentionally unimplemented for now. Keep it *unregistered* so baseline ribbon disabling
  // disables the menu item automatically (and we don't need exemptions).
  assert.doesNotMatch(
    builtins,
    /\bregisterBuiltinCommand\(\s*["']data\.forecast\.whatIfAnalysis\.dataTable["']/,
    "Did not expect Data Table to be registered (unimplemented placeholder)",
  );
  assert.doesNotMatch(
    disabling,
    /["']data\.forecast\.whatIfAnalysis\.dataTable["']/,
    "Did not expect Data Table to be exempted (unimplemented placeholder)",
  );

  // Sanity check: ribbon should be mounted through the CommandRegistry bridge.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
});

