import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("comment tooltip border uses grid line token", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "comments.css");
  const css = fs.readFileSync(cssPath, "utf8");

  assert.match(
    css,
    /\.comment-tooltip\s*\{[\s\S]*?border:\s*1px solid var\(--formula-grid-line\b/,
    "Expected comment tooltip border to use --formula-grid-line so it follows grid theming",
  );
});

