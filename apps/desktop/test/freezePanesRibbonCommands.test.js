import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const CANONICAL_FREEZE_PANES_IDS = [
  "view.freezePanes",
  "view.freezeTopRow",
  "view.freezeFirstColumn",
  "view.unfreezePanes",
];

const LEGACY_FREEZE_PANES_IDS = [
  "view.window.freezePanes.freezePanes",
  "view.window.freezePanes.freezeTopRow",
  "view.window.freezePanes.freezeFirstColumn",
  "view.window.freezePanes.unfreeze",
];

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractObjectLiteralAfter(source, anchorRegex) {
  const match = source.match(anchorRegex);
  assert.ok(match, `Expected to find anchor ${anchorRegex} in source`);
  const anchorIndex = match.index ?? -1;
  assert.ok(anchorIndex >= 0, `Expected match.index for ${anchorRegex}`);

  const braceStart = source.indexOf("{", anchorIndex);
  assert.ok(braceStart !== -1, `Expected to find '{' after ${anchorRegex}`);

  let depth = 0;
  let inString = null;
  let inLineComment = false;
  let inBlockComment = false;

  for (let i = braceStart; i < source.length; i += 1) {
    const char = source[i];
    const next = source[i + 1];

    if (inLineComment) {
      if (char === "\n") inLineComment = false;
      continue;
    }

    if (inBlockComment) {
      if (char === "*" && next === "/") {
        inBlockComment = false;
        i += 1;
      }
      continue;
    }

    if (inString) {
      if (char === "\\") {
        i += 1;
        continue;
      }
      if (char === inString) {
        inString = null;
      }
      continue;
    }

    if (char === "/" && next === "/") {
      inLineComment = true;
      i += 1;
      continue;
    }

    if (char === "/" && next === "*") {
      inBlockComment = true;
      i += 1;
      continue;
    }

    if (char === "'" || char === '"' || char === "`") {
      inString = char;
      continue;
    }

    if (char === "{") {
      depth += 1;
      continue;
    }

    if (char === "}") {
      depth -= 1;
      if (depth === 0) {
        return source.slice(braceStart, i + 1);
      }
    }
  }

  assert.fail(`Expected to find matching '}' for object literal after ${anchorRegex}`);
}

function readRibbonSchemaSource() {
  const schemaDir = path.join(__dirname, "..", "src", "ribbon", "schema");
  try {
    const files = fs
      .readdirSync(schemaDir, { withFileTypes: true })
      .filter((entry) => entry.isFile() && entry.name.endsWith(".ts"))
      .map((entry) => entry.name)
      .sort();
    return files.map((file) => fs.readFileSync(path.join(schemaDir, file), "utf8")).join("\n");
  } catch {
    // Back-compat: older versions kept all tab definitions in ribbonSchema.ts.
    const schemaPath = path.join(__dirname, "..", "src", "ribbon", "ribbonSchema.ts");
    return fs.readFileSync(schemaPath, "utf8");
  }
}

test("Ribbon schema uses canonical View → Freeze Panes command ids", () => {
  const schema = readRibbonSchemaSource();

  // Dropdown trigger id (menu opener).
  assert.match(schema, /\bid:\s*["']view\.window\.freezePanes["']/);

  // Menu items should be the canonical CommandRegistry ids.
  for (const id of CANONICAL_FREEZE_PANES_IDS) {
    assert.match(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbon schema to include ${id}`);
  }

  // Guardrail: do not regress to the old hierarchical ids.
  for (const id of LEGACY_FREEZE_PANES_IDS) {
    assert.doesNotMatch(schema, new RegExp(`\\bid:\\s*["']${escapeRegExp(id)}["']`), `Expected ribbon schema to not include legacy id ${id}`);
  }
});

test("Desktop main.ts does not handle legacy Freeze Panes ribbon ids directly", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  // Legacy ribbon-only ids should not exist anywhere in main.ts.
  for (const id of LEGACY_FREEZE_PANES_IDS) {
    assert.doesNotMatch(main, new RegExp(escapeRegExp(id)), `Expected main.ts to not mention legacy id ${id}`);
  }

  // Ribbon Freeze Panes actions should be executed through CommandRegistry so command-palette recents tracking sees them.
  assert.doesNotMatch(main, /\bapp\.freezePanes\(/, "Expected ribbon Freeze Panes actions to not call app.freezePanes() directly in main.ts");
  assert.doesNotMatch(main, /\bapp\.freezeTopRow\(/, "Expected ribbon Freeze Top Row action to not call app.freezeTopRow() directly in main.ts");
  assert.doesNotMatch(
    main,
    /\bapp\.freezeFirstColumn\(/,
    "Expected ribbon Freeze First Column action to not call app.freezeFirstColumn() directly in main.ts",
  );
  assert.doesNotMatch(main, /\bapp\.unfreezePanes\(/, "Expected ribbon Unfreeze Panes action to not call app.unfreezePanes() directly in main.ts");
});

test("Desktop main.ts does not disable Freeze Panes commands in read-only mode", () => {
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const main = fs.readFileSync(mainPath, "utf8");

  const readOnlyOverrides = extractObjectLiteralAfter(main, /\.\.\.\(isReadOnly\b/);
  const ids = ["view.window.freezePanes", ...CANONICAL_FREEZE_PANES_IDS];
  for (const id of ids) {
    const pattern = new RegExp(`[\"']${escapeRegExp(id)}[\"']\\s*:\\s*true`);
    assert.doesNotMatch(readOnlyOverrides, pattern, `Expected read-only ribbon overrides to not disable ${id}`);
  }
});

test("Builtin command catalog exposes canonical Freeze Panes ids (and does not register the dropdown trigger id)", () => {
  const commandsPath = path.join(__dirname, "..", "src", "commands", "registerBuiltinCommands.ts");
  const commands = fs.readFileSync(commandsPath, "utf8");

  for (const id of CANONICAL_FREEZE_PANES_IDS) {
    const pattern = new RegExp(`\\bregisterBuiltinCommand\\(\\s*["']${escapeRegExp(id)}["']`);
    assert.match(commands, pattern, `Expected builtin command '${id}' to be registered in registerBuiltinCommands.ts`);
  }

  // `view.window.freezePanes` is a ribbon dropdown trigger id (menu opener). It should not be a real
  // command, otherwise the command palette would show duplicate "Freeze Panes" entries.
  assert.doesNotMatch(commands, /\bregisterBuiltinCommand\(\s*["']view\.window\.freezePanes["']/);
});
