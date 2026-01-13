import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema uses canonical Home → Debug command ids (Auditing / Split view / Freeze)", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "schema", "homeTab.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

  const cases = [
    { id: "audit.togglePrecedents", testId: "audit-precedents" },
    { id: "audit.toggleDependents", testId: "audit-dependents" },
    { id: "audit.toggleTransitive", testId: "audit-transitive" },
    { id: "view.splitVertical", testId: "split-vertical" },
    { id: "view.splitHorizontal", testId: "split-horizontal" },
    { id: "view.splitNone", testId: "split-none" },
    { id: "view.freezePanes", testId: "freeze-panes" },
    { id: "view.freezeTopRow", testId: "freeze-top-row" },
    { id: "view.freezeFirstColumn", testId: "freeze-first-column" },
    { id: "view.unfreezePanes", testId: "unfreeze-panes" },
  ];

  for (const { id, testId } of cases) {
    const pattern = new RegExp(
      `\\{[^}]*\\bid:\\s*["']${escapeRegExp(id)}["'][^}]*\\btestId:\\s*["']${escapeRegExp(testId)}["'][^}]*\\}`,
      "m",
    );
    assert.match(schema, pattern, `Expected homeTab.ts to include ${id} with testId ${testId}`);
  }

  // Guardrail: we should not regress back to the legacy ribbon-only ids.
  for (const legacy of [
    "audit-precedents",
    "audit-dependents",
    "audit-transitive",
    "split-vertical",
    "split-horizontal",
    "split-none",
    "freeze-panes",
    "freeze-top-row",
    "freeze-first-column",
    "unfreeze-panes",
  ]) {
    assert.doesNotMatch(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(legacy)}["']`));
  }
});

test("Desktop main.ts no longer handles legacy Home debug ids directly", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  for (const legacy of [
    "audit-precedents",
    "audit-dependents",
    "audit-transitive",
    "split-vertical",
    "split-horizontal",
    "split-none",
    "freeze-panes",
    "freeze-top-row",
    "freeze-first-column",
    "unfreeze-panes",
  ]) {
    assert.doesNotMatch(main, new RegExp(`\\bcase\\s+["']${escapeRegExp(legacy)}["']:`));
  }
});

test("Builtin commands exist for the Home debug actions", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  // Transitive auditing toggle should exist + restore focus.
  assert.match(
    commands,
    /\bregisterBuiltinCommand\(\s*["']audit\.toggleTransitive["'][\s\S]*?app\.toggleAuditingTransitive\(\);[\s\S]*?app\.focus\(\);/m,
    "Expected audit.toggleTransitive to call app.toggleAuditingTransitive() and app.focus()",
  );

  // Explicit split direction commands should exist + restore focus.
  const splitCases = [
    { id: "view.splitVertical", dir: "vertical" },
    { id: "view.splitHorizontal", dir: "horizontal" },
    { id: "view.splitNone", dir: "none" },
  ];
  for (const { id, dir } of splitCases) {
    const pattern = new RegExp(
      `\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["'][\\s\\S]*?layoutController\\.setSplitDirection\\(\\s*["']${escapeRegExp(
        dir,
      )}["']\\s*,\\s*0\\.5\\s*\\);[\\s\\S]*?app\\.focus\\(\\);`,
      "m",
    );
    assert.match(commands, pattern, `Expected ${id} to call layoutController.setSplitDirection("${dir}", 0.5) and app.focus()`);
  }
});

