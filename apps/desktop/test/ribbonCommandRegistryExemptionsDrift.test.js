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

  const registerRe = /\bregisterBuiltinCommand\s*\(\s*["']([^"']+)["']/g;

  for (const file of files) {
    let text = "";
    try {
      text = fs.readFileSync(file, "utf8");
    } catch {
      continue;
    }
    for (const match of text.matchAll(registerRe)) {
      ids.add(match[1]);
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

