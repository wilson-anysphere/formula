import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema includes Data → Queries & Connections controls", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "schema", "dataTab.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

  // Toggle button.
  assert.match(schema, /\bid:\s*["']data\.queriesConnections\.queriesConnections["']/);
  assert.match(schema, /\bkind:\s*["']toggle["']/);

  // Refresh All dropdown + key menu items.
  assert.match(schema, /\bid:\s*["']data\.queriesConnections\.refreshAll["']/);
  assert.match(schema, /\bkind:\s*["']dropdown["']/);

  const refreshMenuIds = [
    "data.queriesConnections.refreshAll",
    "data.queriesConnections.refreshAll.refresh",
    "data.queriesConnections.refreshAll.refreshAllConnections",
    "data.queriesConnections.refreshAll.refreshAllQueries",
  ];
  for (const id of refreshMenuIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`));
  }
});

test("Desktop main.ts wires Data → Queries & Connections controls to the panel + refresh", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Toggle opens/closes the DATA_QUERIES panel via ribbon toggle overrides.
  assert.match(main, /\btoggleOverrides:\s*\{[\s\S]*?["']data\.queriesConnections\.queriesConnections["']\s*:/m);
  assert.match(main, /\bPanelIds\.DATA_QUERIES\b/);
  assert.match(main, /\bopenPanel\(PanelIds\.DATA_QUERIES\)/);
  assert.match(main, /\bclosePanel\(PanelIds\.DATA_QUERIES\)/);

  // Pressed state syncs from layout placement.
  assert.match(
    main,
    /"data\.queriesConnections\.queriesConnections":\s*isPanelOpen\(\s*PanelIds\.DATA_QUERIES\s*\)/,
    "Expected ribbon pressed state to reflect whether the Data Queries panel is open",
  );

  // Refresh All wires to powerQueryService.refreshAll().
  const refreshCommandIds = [
    "data.queriesConnections.refreshAll",
    "data.queriesConnections.refreshAll.refresh",
    "data.queriesConnections.refreshAll.refreshAllConnections",
    "data.queriesConnections.refreshAll.refreshAllQueries",
  ];
  for (const id of refreshCommandIds) {
    assert.match(
      main,
      new RegExp(`\\bcommandId\\s*===\\s*["']${escapeRegExp(id)}["']`),
      `Expected main.ts to handle refresh command id ${id}`,
    );
  }
  assert.match(main, /\brefreshAll\(\)/);
});
