import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function findButtonBlockByTestId(source, testId) {
  const re = new RegExp(
    // The Home debug panel buttons are simple object literals with no nested `{}`.
    // Keep the matcher strict so we assert `id` + `testId` belong to the same button.
    `\\{[^{}]*\\btestId:\\s*["']${escapeRegExp(testId)}["'][^{}]*\\}`,
    "m",
  );
  const match = source.match(re);
  assert.ok(match, `Expected ribbon schema to include a button with testId ${testId}`);
  return match[0];
}

test("Home → Debug → Panels ribbon buttons use canonical CommandRegistry ids (stable testIds)", () => {
  const homeTabPath = path.join(__dirname, "..", "src", "ribbon", "schema", "homeTab.ts");
  const homeTab = fs.readFileSync(homeTabPath, "utf8");

  const expected = [
    { commandId: "view.togglePanel.aiAudit", testId: "open-panel-ai-audit" },
    // Legacy UI hook kept for Playwright stability.
    { commandId: "view.togglePanel.aiAudit", testId: "open-ai-audit-panel" },
    { commandId: "view.togglePanel.dataQueries", testId: "open-data-queries-panel" },
    { commandId: "view.togglePanel.macros", testId: "open-macros-panel" },
    { commandId: "view.togglePanel.scriptEditor", testId: "open-script-editor-panel" },
    { commandId: "view.togglePanel.python", testId: "open-python-panel" },
    { commandId: "view.togglePanel.extensions", testId: "open-extensions-panel" },
    { commandId: "view.togglePanel.vbaMigrate", testId: "open-vba-migrate-panel" },
    { commandId: "comments.togglePanel", testId: "open-comments-panel" },
  ];

  for (const { commandId, testId } of expected) {
    const block = findButtonBlockByTestId(homeTab, testId);
    assert.match(
      block,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(commandId)}["']`),
      `Expected Home tab debug panel button ${testId} to use command id ${commandId}`,
    );
  }

  // Ensure we don't regress to legacy ribbon-only "open-*" command ids.
  for (const legacyId of [
    "open-panel-ai-audit",
    "open-ai-audit-panel",
    "open-data-queries-panel",
    "open-macros-panel",
    "open-script-editor-panel",
    "open-python-panel",
    "open-extensions-panel",
    "open-vba-migrate-panel",
    "open-comments-panel",
  ]) {
    assert.doesNotMatch(
      homeTab,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(legacyId)}["']`),
      `Did not expect Home tab schema to expose legacy command id ${legacyId}`,
    );
  }
});

test("Home → Debug → Panels commands are registered in CommandRegistry (not wired only in main.ts)", () => {
  const builtinsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const builtins = fs.readFileSync(builtinsPath, "utf8");

  // Ensure panel toggles are real built-in commands so the ribbon shares wiring with the
  // command palette / keybindings and recents tracking.
  for (const commandId of [
    "view.togglePanel.aiAudit",
    "view.togglePanel.dataQueries",
    "view.togglePanel.macros",
    "view.togglePanel.scriptEditor",
    "view.togglePanel.python",
    "view.togglePanel.extensions",
    "view.togglePanel.vbaMigrate",
    "comments.togglePanel",
  ]) {
    assert.match(
      builtins,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*\\n\\s*["']${escapeRegExp(commandId)}["']`),
      `Expected registerBuiltinCommands.ts to register ${commandId}`,
    );
  }

  // Spot-check the VBA migrate toggle wiring (added during this migration).
  const vbaBlock = builtins.match(/registerBuiltinCommand\(\s*\n\s*["']view\.togglePanel\.vbaMigrate["'][\s\S]*?\n\s*\);\n/s);
  assert.ok(vbaBlock, "Expected to find view.togglePanel.vbaMigrate registration block");
  assert.match(
    vbaBlock[0],
    /\btoggleDockPanel\(PanelIds\.VBA_MIGRATE\)/,
    "Expected view.togglePanel.vbaMigrate to toggle PanelIds.VBA_MIGRATE via toggleDockPanel",
  );

  const commentsBlock = builtins.match(/registerBuiltinCommand\(\s*\n\s*["']comments\.togglePanel["'][\s\S]*?\n\s*\);\n/s);
  assert.ok(commentsBlock, "Expected to find comments.togglePanel registration block");
  assert.match(
    commentsBlock[0],
    /\(\)\s*=>\s*app\.toggleCommentsPanel\(\)/,
    "Expected comments.togglePanel to call app.toggleCommentsPanel()",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Ensure we don't accidentally re-introduce ribbon-only open-* cases (these should route via CommandRegistry).
  for (const legacy of [
    "open-panel-ai-audit",
    "open-ai-audit-panel",
    "open-data-queries-panel",
    "open-macros-panel",
    "open-script-editor-panel",
    "open-python-panel",
    "open-extensions-panel",
    "open-vba-migrate-panel",
    "open-comments-panel",
  ]) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(legacy)}["']`),
      `Did not expect main.ts to handle ribbon-only command case ${legacy}`,
    );
  }

  // Pressed state sync should reflect whether the panels are open so the debug buttons
  // visually track the current layout state (Excel-style toggle behavior).
  const expectedPressedMappings = [
    { commandId: "view.togglePanel.aiAudit", panelId: "AI_AUDIT" },
    { commandId: "view.togglePanel.dataQueries", panelId: "DATA_QUERIES" },
    { commandId: "view.togglePanel.macros", panelId: "MACROS" },
    { commandId: "view.togglePanel.scriptEditor", panelId: "SCRIPT_EDITOR" },
    { commandId: "view.togglePanel.python", panelId: "PYTHON" },
    { commandId: "view.togglePanel.extensions", panelId: "EXTENSIONS" },
    { commandId: "view.togglePanel.vbaMigrate", panelId: "VBA_MIGRATE" },
  ];
  for (const { commandId, panelId } of expectedPressedMappings) {
    assert.match(
      main,
      new RegExp(`["']${escapeRegExp(commandId)}["']:\\s*isPanelOpen\\(\\s*PanelIds\\.${escapeRegExp(panelId)}\\s*\\)`),
      `Expected main.ts to sync pressed state for ${commandId} via PanelIds.${panelId}`,
    );
  }
  assert.match(
    main,
    /["']comments\.togglePanel["']:\s*app\.isCommentsPanelVisible\(\)/,
    "Expected main.ts to sync pressed state for comments.togglePanel",
  );
});
