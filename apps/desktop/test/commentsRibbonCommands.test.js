import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

test("Ribbon schema uses canonical Review → Comments command ids", () => {
  const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
  const schema = fs.readFileSync(schemaPath, "utf8");

  // Review → Comments group should be wired to the stable builtin command ids.
  const commandIds = ["comments.addComment", "comments.togglePanel"];
  for (const id of commandIds) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbonSchema.ts to include ${id}`);
  }

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

  // Ribbon command activation should execute the builtin commands (so invocation is recorded
  // by command palette recents tracking and shares the same path as keybindings).
  assert.match(
    main,
    new RegExp(
      `if\\s*\\(commandId\\s*===\\s*["']comments\\.togglePanel["']\\s*\\|\\|\\s*commandId\\s*===\\s*["']comments\\.addComment["']\\)\\s*\\{[\\s\\S]*?executeBuiltinCommand\\(commandId\\);`,
      "m",
    ),
    "Expected ribbon onCommand handler to execute comments.* via executeBuiltinCommand(commandId)",
  );

  // Guardrail: the legacy review.comments.* ids should not be handled in main.ts anymore.
  assert.doesNotMatch(main, /\breview\.comments\.newComment\b/);
  assert.doesNotMatch(main, /\breview\.comments\.showComments\b/);
});

