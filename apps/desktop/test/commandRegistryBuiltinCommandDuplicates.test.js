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

test("command registration sources do not register the same builtin command id twice within a file", () => {
  const srcRoot = path.join(__dirname, "..", "src");
  const commandsRoot = path.join(srcRoot, "commands");

  /** @type {string[]} */
  const files = [];
  collectSourceFiles(commandsRoot, files);
  files.sort((a, b) => a.localeCompare(b));

  const registerRe = /\bregisterBuiltinCommand\s*\(\s*["']([^"']+)["']/g;
  /** @type {Array<{ file: string, duplicates: Array<{ value: string, count: number }> }>} */
  const duplicatesByFile = [];

  for (const file of files) {
    const text = fs.readFileSync(file, "utf8");
    const ids = [];
    for (const match of text.matchAll(registerRe)) {
      ids.push(match[1]);
    }
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

