import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Sort/Filter ribbon custom sort aliases are hidden from the command palette", () => {
  const sourcePath = path.join(__dirname, "..", "src", "commands", "registerSortFilterCommands.ts");
  const source = fs.readFileSync(sourcePath, "utf8");

  // The Home tab uses a ribbon-scoped id for Custom Sort, but it is an alias of the
  // Data tab command. Ensure it's hidden to avoid duplicate "Custom Sort…" entries
  // in the command palette.
  assert.match(source, /\bregisterCustomSortCommand\(\s*SORT_FILTER_RIBBON_COMMANDS\.homeCustomSort[\s\S]*?\bwhen:\s*["']false["']/);

  // The Data tab command should remain the canonical visible command (i.e. it should
  // not be explicitly hidden).
  assert.doesNotMatch(source, /\bregisterCustomSortCommand\(\s*SORT_FILTER_RIBBON_COMMANDS\.dataCustomSort[\s\S]*?\bwhen:\s*["']false["']/);
});

