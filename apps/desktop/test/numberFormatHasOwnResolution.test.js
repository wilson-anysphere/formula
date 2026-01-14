import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("numberFormat resolution does not fall back to snake_case number_format (null/undefined should override)", () => {
  const srcRoot = path.join(__dirname, "..", "src");

  /** @type {string[]} */
  const files = [];
  collectSourceFiles(srcRoot, files);

  // When `numberFormat` is explicitly set to `null` we want that to override any imported
  // `number_format` value (e.g. XLSX date formats). Using `??` or `||` will incorrectly
  // fall back when the key is present but cleared.
  const legacyNullishFallbackRe = /\bnumberFormat\s*\?\?\s*[^;\n]*\bnumber_format\b/;
  const legacyOrFallbackRe = /\bnumberFormat\s*\|\|\s*[^;\n]*\bnumber_format\b/;

  for (const file of files) {
    const text = stripComments(fs.readFileSync(file, "utf8"));
    const rel = path.relative(srcRoot, file);
    assert.doesNotMatch(
      text,
      legacyNullishFallbackRe,
      `${rel} should use getStyleNumberFormat (hasOwn semantics) instead of \`numberFormat ?? number_format\``,
    );
    assert.doesNotMatch(
      text,
      legacyOrFallbackRe,
      `${rel} should use getStyleNumberFormat (hasOwn semantics) instead of \`numberFormat || number_format\``,
    );
  }
});

/**
 * @param {string} dir
 * @param {string[]} out
 */
function collectSourceFiles(dir, out) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name.startsWith(".")) continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      collectSourceFiles(full, out);
      continue;
    }
    if (!entry.isFile()) continue;
    if (entry.name.endsWith(".d.ts") || entry.name.endsWith(".d.tsx")) continue;
    const ext = path.extname(entry.name);
    if (ext !== ".ts" && ext !== ".tsx" && ext !== ".js" && ext !== ".jsx") continue;
    out.push(full);
  }
}
