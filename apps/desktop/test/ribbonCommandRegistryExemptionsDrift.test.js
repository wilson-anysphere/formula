import fs from "node:fs";
import path from "node:path";
import assert from "node:assert/strict";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { COMMAND_REGISTRY_EXEMPT_IDS } from "../src/ribbon/ribbonCommandRegistryDisabling.js";
import { defaultRibbonSchema } from "../src/ribbon/ribbonSchema.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function collectDefaultRibbonIds() {
  const ids = new Set();
  for (const tab of defaultRibbonSchema.tabs) {
    for (const group of tab.groups) {
      for (const button of group.buttons) {
        ids.add(button.id);
        for (const menuItem of button.menuItems ?? []) {
          ids.add(menuItem.id);
        }
      }
    }
  }
  return ids;
}

function collectSourceFiles(dir, out) {
  let entries = [];
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      collectSourceFiles(fullPath, out);
      continue;
    }
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".ts") && !entry.name.endsWith(".js")) continue;
    if (entry.name.endsWith(".d.ts")) continue;
    out.push(fullPath);
  }
}

function collectRegisteredBuiltinCommandIds() {
  const ids = new Set();

  const srcRoot = path.join(__dirname, "..", "src");
  const files = [];
  collectSourceFiles(path.join(srcRoot, "commands"), files);
  files.push(path.join(srcRoot, "main.ts"));

  const extractStringLiteral = (value) => {
    const trimmed = String(value ?? "").trim();
    if (!trimmed) return null;
    const match = trimmed.match(/^["']([^"']+)["']$/);
    return match ? match[1] : null;
  };

  /**
   * Best-effort resolution for command-id expressions in registration call sites.
   *
   * Supports:
   * - String literals: `"foo.bar"`
   * - Const string bindings: `const ID = "foo.bar";`
   * - Const object properties: `const IDS = { save: "foo.bar" } as const;` -> `IDS.save`
   */
  const resolveCommandIdExpr = (expr, constStrings, constObjects) => {
    const trimmed = String(expr ?? "").trim();
    if (!trimmed) return null;

    const literal = extractStringLiteral(trimmed);
    if (literal) return literal;

    const identMatch = trimmed.match(/^(\w+)$/);
    if (identMatch) {
      return constStrings.get(identMatch[1]) ?? null;
    }

    const propMatch = trimmed.match(/^(\w+)\.(\w+)$/);
    if (propMatch) {
      const obj = constObjects.get(propMatch[1]);
      if (!obj) return null;
      return obj.get(propMatch[2]) ?? null;
    }

    return null;
  };

  const registerCallRe = /\bregisterBuiltinCommand\s*\(\s*([^,]+?)\s*,/g;
  const constStringRe = /\b(?:export\s+)?const\s+(\w+)\s*=\s*["']([^"']+)["']\s*(?:as\s+const)?\s*;/g;
  const constObjectRe = /\b(?:export\s+)?const\s+(\w+)\s*=\s*{([\s\S]*?)}\s*(?:as\s+const)?\s*;/g;
  const constArrayRe = /\b(?:export\s+)?const\s+(\w+)\s*=\s*\[([\s\S]*?)]\s*(?:as\s+const)?\s*;/g;
  const objectPairRe = /\b(\w+)\s*:\s*["']([^"']+)["']/g;

  // Detect patterns like:
  //   for (const commandId of RIBBON_MACRO_COMMAND_IDS) {
  //     commandRegistry.registerBuiltinCommand(commandId, ...)
  //   }
  const loopRegistrationRe = /\bfor\s*\(\s*const\s+(\w+)\s+of\s+(\w+)\s*\)\s*{[\s\S]*?registerBuiltinCommand\s*\(\s*\1\b/g;

  // Detect helper functions that register commands via their first argument:
  //   const registerSortCommand = (commandId: string, ...) => { registerBuiltinCommand(commandId, ...) }
  // Then collect call sites:
  //   registerSortCommand("foo.bar", ...)
  //   registerSortCommand(SOME_IDS.foo, ...)
  const arrowHelperDefRe = /\bconst\s+(\w+)\s*=\s*\(\s*(\w+)(?:\s*:\s*[^,\)]+)?/g;
  const functionHelperDefRe = /\bfunction\s+(\w+)\s*\(\s*(\w+)(?:\s*:\s*[^,\)]+)?/g;

  for (const file of files) {
    let text = "";
    try {
      text = fs.readFileSync(file, "utf8");
    } catch {
      continue;
    }

    /** @type {Map<string, string>} */
    const constStrings = new Map();
    for (const match of text.matchAll(constStringRe)) {
      constStrings.set(match[1], match[2]);
    }

    /** @type {Map<string, Map<string, string>>} */
    const constObjects = new Map();
    for (const match of text.matchAll(constObjectRe)) {
      const name = match[1];
      const body = match[2] ?? "";
      const pairs = new Map();
      for (const pair of body.matchAll(objectPairRe)) {
        pairs.set(pair[1], pair[2]);
      }
      if (pairs.size > 0) constObjects.set(name, pairs);
    }

    /** @type {Map<string, string[]>} */
    const constArrays = new Map();
    for (const match of text.matchAll(constArrayRe)) {
      const name = match[1];
      const body = match[2] ?? "";
      const items = [];
      for (const item of body.matchAll(/["']([^"']+)["']/g)) {
        items.push(item[1]);
      }
      if (items.length > 0) constArrays.set(name, items);
    }

    // 1) Direct registerBuiltinCommand(...) calls.
    for (const match of text.matchAll(registerCallRe)) {
      const expr = match[1];
      const id = resolveCommandIdExpr(expr, constStrings, constObjects);
      if (id) ids.add(id);
    }

    // 2) Loop-based registrations over const string arrays.
    for (const match of text.matchAll(loopRegistrationRe)) {
      const arrayName = match[2];
      const items = constArrays.get(arrayName);
      if (!items) continue;
      for (const id of items) ids.add(id);
    }

    // 3) Helper function registration wrappers.
    /** @type {Set<string>} */
    const helperNames = new Set();
    const detectHelper = (name, param, startIndex) => {
      // Keep the scan local-ish to reduce false matches.
      const start = typeof startIndex === "number" && Number.isFinite(startIndex) ? Math.max(0, startIndex) : 0;
      const window = text.slice(start, start + 8_000);
      const paramRe = new RegExp(`\\bregisterBuiltinCommand\\s*\\(\\s*${param}\\b`);
      if (paramRe.test(window)) helperNames.add(name);
    };

    for (const match of text.matchAll(arrowHelperDefRe)) {
      detectHelper(match[1], match[2], match.index);
    }
    for (const match of text.matchAll(functionHelperDefRe)) {
      detectHelper(match[1], match[2], match.index);
    }

    for (const helperName of helperNames) {
      const callRe = new RegExp(`\\b${helperName}\\s*\\(\\s*([^,\\)]+)`, "g");
      for (const call of text.matchAll(callRe)) {
        const id = resolveCommandIdExpr(call[1], constStrings, constObjects);
        if (id) ids.add(id);
      }
    }
  }

  return ids;
}

test("COMMAND_REGISTRY_EXEMPT_IDS stays in sync with defaultRibbonSchema", () => {
  const ribbonIds = collectDefaultRibbonIds();
  const staleExemptions = [...COMMAND_REGISTRY_EXEMPT_IDS].filter((id) => !ribbonIds.has(id)).sort();

  assert.deepEqual(
    staleExemptions,
    [],
    `Exemptions contain ids that are no longer present in defaultRibbonSchema:\n${staleExemptions.map((id) => `- ${id}`).join("\n")}`,
  );
});

test("COMMAND_REGISTRY_EXEMPT_IDS does not contain registered builtin CommandRegistry ids", () => {
  const registered = collectRegisteredBuiltinCommandIds();
  const implementedExemptions = [...COMMAND_REGISTRY_EXEMPT_IDS].filter((id) => registered.has(id)).sort();

  assert.deepEqual(
    implementedExemptions,
    [],
    [
      "Exemptions contain ids that appear to be registered via registerBuiltinCommand(...) calls (please remove them from COMMAND_REGISTRY_EXEMPT_IDS):",
      ...implementedExemptions.map((id) => `- ${id}`),
    ].join("\n"),
  );
});
