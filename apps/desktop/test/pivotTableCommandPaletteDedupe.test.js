import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("PivotTable ribbon alias is hidden from the command palette", () => {
  const sourcePath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const source = stripComments(fs.readFileSync(sourcePath, "utf8"));

  function escapeRegExp(value) {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  const extractRegistration = (commandId) => {
    const re = new RegExp(String.raw`\bregisterBuiltinCommand\(\s*["']${commandId}["'][\s\S]*?\n\s*\);\n`);
    return source.match(re)?.[0] ?? null;
  };

  const alias = extractRegistration(escapeRegExp("insert.tables.pivotTable.fromTableRange"));
  assert.ok(alias, "Expected registerBuiltinCommands.ts to register insert.tables.pivotTable.fromTableRange");
  assert.match(
    alias,
    /\bwhen:\s*["']false["']/,
    "Expected insert.tables.pivotTable.fromTableRange to be hidden via when: \"false\" (avoid duplicate command palette entries)",
  );

  const canonical = extractRegistration(escapeRegExp("view.insertPivotTable"));
  assert.ok(canonical, "Expected registerBuiltinCommands.ts to register view.insertPivotTable");
  assert.doesNotMatch(
    canonical,
    /\bwhen:\s*["']false["']/,
    "Expected view.insertPivotTable to remain visible in the command palette (canonical command)",
  );
});
