import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function walk(dirPath) {
  /** @type {string[]} */
  const entries = [];
  for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      entries.push(...walk(fullPath));
    } else {
      entries.push(fullPath);
    }
  }
  return entries;
}

test("core UI does not hardcode colors outside tokens.css", () => {
  const srcRoot = path.join(__dirname, "..", "src");
  const files = walk(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    if (rel === "styles/tokens.css") return false;
    if (rel.startsWith("grid/presence-renderer/")) return false;
    // Rich text styles include data-driven colors (cell formatting).
    if (rel.startsWith("grid/text/rich-text/")) return false;
    // Conditional formatting colors are data-driven (cell formatting).
    if (rel.startsWith("grid/conditional-formatting/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    if (/\.(test|spec)\.[jt]sx?$/.test(rel)) return false;
    return (
      rel.endsWith(".css") ||
      rel.endsWith(".js") ||
      rel.endsWith(".ts") ||
      rel.endsWith(".tsx")
    );
  });

  const hexColor = /#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})\b/gi;
  // Only flag *hardcoded* rgb/rgba literals (e.g. `rgb(0,0,0)`), not references
  // in parsing utilities (regexes) or template strings.
  const rgbColor = /\brgb(a)?\s*\(\s*[\d.]/gi;

  /** @type {{ file: string, match: string }[]} */
  const violations = [];

  for (const file of files) {
    const content = fs.readFileSync(file, "utf8");
    const hex = content.match(hexColor);
    const rgb = content.match(rgbColor);
    if (hex) violations.push({ file, match: hex[0] });
    if (rgb) violations.push({ file, match: "rgb(...)" });
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded colors in core UI files:\\n${violations
      .map((v) => `- ${path.relative(srcRoot, v.file)}: ${v.match}`)
      .join("\\n")}`
  );
});
