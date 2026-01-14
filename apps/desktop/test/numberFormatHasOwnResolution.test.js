import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("numberFormat resolution does not use nullish fallback to snake_case number_format", () => {
  const files = [
    {
      name: "SpreadsheetApp",
      file: path.join(__dirname, "..", "src", "app", "spreadsheetApp.ts"),
    },
    {
      name: "desktop main.ts",
      file: path.join(__dirname, "..", "src", "main.ts"),
    },
  ];

  // When `numberFormat` is explicitly set to `null` we want that to override any imported
  // `number_format` value (e.g. XLSX date formats). Using `??` will incorrectly fall back.
  const legacyFallbackRe = /\bnumberFormat\s*\?\?\s*[^;\n]*\bnumber_format\b/;

  for (const { name, file } of files) {
    const text = fs.readFileSync(file, "utf8");
    assert.doesNotMatch(
      text,
      legacyFallbackRe,
      `${name} should use getStyleNumberFormat (hasOwn semantics) instead of \`numberFormat ?? number_format\``,
    );
  }
});

