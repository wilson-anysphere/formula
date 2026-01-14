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
  const main = fs.readFileSync(mainPath, "utf8");

  // Pressed state should follow the SpreadsheetApp comments panel visibility.
  assert.match(
    main,
    /"comments\.togglePanel":\s*app\.isCommentsPanelVisible\(\)/,
    "Expected main.ts to sync ribbon pressed state from app.isCommentsPanelVisible()",
  );

  // Comments commands are registered as builtin commands, and are dispatched via
  // `createRibbonActionsFromCommands` (so they share the same execution path as command palette / keybindings).
  assert.doesNotMatch(
    main,
    /\bcase\s+["']comments\.togglePanel["']:/,
    "Expected main.ts to not handle comments.togglePanel via switch case",
  );
  assert.doesNotMatch(
    main,
    /\bcase\s+["']comments\.addComment["']:/,
    "Expected main.ts to not handle comments.addComment via switch case",
  );

  // Guardrail: the legacy review.comments.* ids should not be handled in main.ts anymore.
  assert.doesNotMatch(main, /\breview\.comments\.newComment\b/);
  assert.doesNotMatch(main, /\breview\.comments\.showComments\b/);
});

test("Builtin Comments commands are registered with the expected behavior", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  // Toggle command: best-effort toggle semantics.
  assert.match(
    commands,
    /\bregisterBuiltinCommand\(\s*["']comments\.togglePanel["'][\s\S]*?=>\s*app\.toggleCommentsPanel\(\)/,
    "Expected comments.togglePanel to dispatch to app.toggleCommentsPanel()",
  );

  // Add comment command: must open the panel and focus the input (Shift+F2 behavior).
  assert.match(
    commands,
    /\bregisterBuiltinCommand\(\s*["']comments\.addComment["'][\s\S]*?if\s*\(app\.isEditing\(\)\)\s*return;[\s\S]*?app\.openCommentsPanel\(\);[\s\S]*?app\.focusNewCommentInput\(\);/,
    "Expected comments.addComment to open comments panel + focus new comment input (guarded by app.isEditing())",
  );
});
