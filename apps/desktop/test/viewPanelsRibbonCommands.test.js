import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("View → Panels ribbon buttons use canonical CommandRegistry ids (stable testIds)", () => {
  const viewTabPath = path.join(__dirname, "..", "src", "ribbon", "schema", "viewTab.ts");
  const viewTab = fs.readFileSync(viewTabPath, "utf8");

  const expected = [
    { commandId: "view.togglePanel.marketplace", testId: "open-marketplace-panel" },
    { commandId: "view.togglePanel.versionHistory", testId: "open-version-history-panel" },
    { commandId: "view.togglePanel.branchManager", testId: "open-branch-manager-panel" },
  ];

  for (const { commandId, testId } of expected) {
    assert.match(
      viewTab,
      new RegExp(`\\bid:\\s*["']${escapeRegExp(commandId)}["']`),
      `Expected View tab schema to include button id ${commandId}`,
    );
    assert.match(
      viewTab,
      new RegExp(`\\btestId:\\s*["']${escapeRegExp(testId)}["']`),
      `Expected View tab schema to include testId ${testId}`,
    );
  }
});

test("View → Panels ribbon actions are wired through CommandRegistry (no ribbon-only open-* switch cases)", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  // Ensure the canonical commands exist.
  for (const id of ["view.togglePanel.marketplace", "view.togglePanel.versionHistory", "view.togglePanel.branchManager"]) {
    assert.match(
      commands,
      new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`),
      `Expected registerBuiltinCommands.ts to register ${id}`,
    );
  }

  // Marketplace should remain usable without eagerly starting the extension host.
  // (Playwright asserts this explicitly in marketplace-panel.spec.ts.)
  const marketplaceBlock = commands.match(
    /registerBuiltinCommand\(\s*["']view\.togglePanel\.marketplace["'][\s\S]*?\n\s*\);\n/s,
  );
  assert.ok(marketplaceBlock, "Expected to find view.togglePanel.marketplace registration block");
  assert.match(
    marketplaceBlock[0],
    /\btoggleDockPanel\(PanelIds\.MARKETPLACE\)/,
    "Expected Marketplace panel toggle to call toggleDockPanel(PanelIds.MARKETPLACE)",
  );
  assert.doesNotMatch(
    marketplaceBlock[0],
    /\bensureExtensionsLoaded\b/,
    "Marketplace toggle should not call ensureExtensionsLoaded (keep extension host lazy)",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  for (const legacy of ["open-marketplace-panel", "open-version-history-panel", "open-branch-manager-panel"]) {
    assert.doesNotMatch(
      main,
      new RegExp(`\\bcase\\s+["']${escapeRegExp(legacy)}["']`),
      `Did not expect main.ts to handle ribbon-only command case ${legacy}`,
    );
  }
});

