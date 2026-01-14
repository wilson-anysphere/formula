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

test("Ribbon schema uses canonical Review → Comments command ids", () => {
  const schema = readRibbonSchemaSource("reviewTab.ts");

  // Review → Comments group should be wired to the stable builtin command ids.
  const commandIds = ["comments.addComment", "comments.togglePanel"];
  for (const id of commandIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected reviewTab.ts to include ${id}`);
  }

  // Preserve key metadata so the UI stays stable.
  assert.match(
    schema,
    /\bid:\s*["']comments\.addComment["'][\s\S]*?\blabel:\s*["']New Comment["'][\s\S]*?\bariaLabel:\s*["']New Comment["'][\s\S]*?\biconId:\s*["']comment["'][\s\S]*?\bsize:\s*["']large["']/,
    "Expected comments.addComment ribbon button to preserve label/ariaLabel/iconId/size",
  );
  assert.match(
    schema,
    /\bid:\s*["']comments\.togglePanel["'][\s\S]*?\blabel:\s*["']Show Comments["'][\s\S]*?\bariaLabel:\s*["']Show Comments["'][\s\S]*?\biconId:\s*["']eye["'][\s\S]*?\bkind:\s*["']toggle["']/,
    "Expected comments.togglePanel ribbon button to preserve label/ariaLabel/iconId/kind",
  );

  // Guardrail: we should not regress back to legacy ribbon-only ids.
  assert.doesNotMatch(schema, /\bid:\s*["']review\.comments\.newComment["']/);
  assert.doesNotMatch(schema, /\bid:\s*["']review\.comments\.showComments["']/);
});

test("Desktop main.ts syncs Comments pressed state + dispatches via CommandRegistry", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = stripComments(fs.readFileSync(mainPath, "utf8"));
  const routerPath = path.join(__dirname, "..", "src", "ribbon", "ribbonCommandRouter.ts");
  const router = stripComments(fs.readFileSync(routerPath, "utf8"));

  // Pressed state should follow the SpreadsheetApp comments panel visibility.
  assert.match(
    main,
    /"comments\.togglePanel":\s*app\.isCommentsPanelVisible\(\)/,
    "Expected main.ts to sync ribbon pressed state from app.isCommentsPanelVisible()",
  );

  // Ribbon command activation should execute registered commands via the CommandRegistry
  // bridge (via the ribbon command router). Avoid bespoke `handleRibbonCommand` routing for
  // comments.* ids so command palette recents + keybindings share the same path.
  assert.match(main, /\bcreateRibbonActions\(/);
  assert.match(router, /\bcreateRibbonActionsFromCommands\(/);
  assert.doesNotMatch(
    router,
    /\btoggleOverrides\s*(?:(?::\s*(?:[^=]|=(?!\s*\{))+\s*)?=|:)\s*\{[\s\S]*?["']comments\.togglePanel["']\s*:/m,
  );
  assert.doesNotMatch(
    router,
    /\bcommandOverrides\s*(?:(?::\s*(?:[^=]|=(?!\s*\{))+\s*)?=|:)\s*\{[\s\S]*?["']comments\.togglePanel["']\s*:/m,
  );
  assert.doesNotMatch(
    router,
    /\bcommandOverrides\s*(?:(?::\s*(?:[^=]|=(?!\s*\{))+\s*)?=|:)\s*\{[\s\S]*?["']comments\.addComment["']\s*:/m,
  );
  assert.doesNotMatch(
    router,
    /\bcommandId\.startsWith\(\s*["']comments\./,
    "Did not expect ribbonCommandRouter.ts to add bespoke comments.* prefix routing (dispatch should go through CommandRegistry)",
  );
  assert.doesNotMatch(
    router,
    /\bcase\s+["']comments\.togglePanel["']:/,
    "Expected ribbonCommandRouter.ts to not handle comments.togglePanel via switch case (should dispatch via CommandRegistry)",
  );
  assert.doesNotMatch(
    router,
    /\bcase\s+["']comments\.addComment["']:/,
    "Expected ribbonCommandRouter.ts to not handle comments.addComment via switch case (should dispatch via CommandRegistry)",
  );
  assert.doesNotMatch(
    router,
    /\bcommandId\s*===\s*["']comments\.togglePanel["']/,
    "Expected ribbonCommandRouter.ts to not special-case comments.togglePanel via commandId === checks (should dispatch via CommandRegistry)",
  );
  assert.doesNotMatch(
    router,
    /\bcommandId\s*===\s*["']comments\.addComment["']/,
    "Expected ribbonCommandRouter.ts to not special-case comments.addComment via commandId === checks (should dispatch via CommandRegistry)",
  );

  // Guardrail: the legacy review.comments.* ids should not be handled in main.ts anymore.
  assert.doesNotMatch(main, /\breview\.comments\.newComment\b/);
  assert.doesNotMatch(main, /\breview\.comments\.showComments\b/);
});

test("Builtin Comments commands are registered with the expected behavior", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const commands = stripComments(fs.readFileSync(commandsPath, "utf8"));

  // Toggle command: best-effort toggle semantics.
  assert.match(
    commands,
    /\bregisterBuiltinCommand\(\s*["']comments\.togglePanel["'][\s\S]*?=>\s*app\.toggleCommentsPanel\(\)/,
    "Expected comments.togglePanel to dispatch to app.toggleCommentsPanel()",
  );

  // Add comment command: must open the panel and focus the input (Shift+F2 behavior).
  assert.match(
    commands,
    /\bregisterBuiltinCommand\(\s*["']comments\.addComment["'][\s\S]*?if\s*\(isEditingFn\(\)\)\s*return;[\s\S]*?app\.openCommentsPanel\(\);[\s\S]*?app\.focusNewCommentInput\(\);/,
    "Expected comments.addComment to open comments panel + focus new comment input (guarded by isEditingFn())",
  );
});
