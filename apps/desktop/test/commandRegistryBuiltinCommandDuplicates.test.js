import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

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

function findDuplicates(values) {
  const counts = new Map();
  for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1);
  return [...counts.entries()]
    .filter(([, count]) => count > 1)
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value));
}

function collectRegisteredIdsFromSource(source) {
  const registerRe = /\bregisterBuiltinCommand\s*\(\s*["']([^"']+)["']/g;
  const ids = [];
  for (const match of source.matchAll(registerRe)) {
    ids.push(match[1]);
  }
  return ids;
}

test("command registration sources do not register the same builtin command id twice within a file", () => {
  const srcRoot = path.join(__dirname, "..", "src");
  const commandsRoot = path.join(srcRoot, "commands");

  /** @type {string[]} */
  const files = [];
  collectSourceFiles(commandsRoot, files);
  files.sort((a, b) => a.localeCompare(b));

  /** @type {Array<{ file: string, duplicates: Array<{ value: string, count: number }> }>} */
  const duplicatesByFile = [];

  for (const file of files) {
    const text = fs.readFileSync(file, "utf8");
    const ids = collectRegisteredIdsFromSource(text);
    const duplicates = findDuplicates(ids);
    if (duplicates.length > 0) {
      duplicatesByFile.push({ file: path.relative(srcRoot, file), duplicates });
    }
  }

  assert.deepEqual(
    duplicatesByFile,
    [],
    `Found duplicate builtin command registrations within a single file:\n${duplicatesByFile
      .map(
        ({ file, duplicates }) =>
          `- ${file}\n${duplicates.map(({ value, count }) => `  - ${value} (${count}x)`).join("\n")}`,
      )
      .join("\n")}`,
  );
});

test("main.ts does not re-register builtin command ids already registered by src/commands", () => {
  const srcRoot = path.join(__dirname, "..", "src");
  const commandsRoot = path.join(srcRoot, "commands");
  const mainPath = path.join(srcRoot, "main.ts");

  /** @type {string[]} */
  const commandFiles = [];
  collectSourceFiles(commandsRoot, commandFiles);
  commandFiles.sort((a, b) => a.localeCompare(b));

  const commandIds = new Set();
  for (const file of commandFiles) {
    const text = fs.readFileSync(file, "utf8");
    for (const id of collectRegisteredIdsFromSource(text)) {
      commandIds.add(id);
    }
  }

  const mainText = fs.readFileSync(mainPath, "utf8");
  const overlaps = collectRegisteredIdsFromSource(mainText).filter((id) => commandIds.has(id));
  overlaps.sort((a, b) => a.localeCompare(b));

  assert.deepEqual(
    overlaps,
    [],
    `Found builtin command ids registered in both src/main.ts and src/commands/*:\n${overlaps.map((id) => `- ${id}`).join("\n")}`,
  );
});

test("main.ts does not register the same builtin command id twice", () => {
  const srcRoot = path.join(__dirname, "..", "src");
  const mainPath = path.join(srcRoot, "main.ts");
  const mainText = fs.readFileSync(mainPath, "utf8");
  const ids = collectRegisteredIdsFromSource(mainText);
  const duplicates = findDuplicates(ids);

  assert.deepEqual(
    duplicates,
    [],
    `src/main.ts contains duplicate registerBuiltinCommand(...) ids:\n${duplicates.map(({ value, count }) => `- ${value} (${count}x)`).join("\n")}`,
  );
});
