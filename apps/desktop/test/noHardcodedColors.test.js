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
    // Charts scene graph utilities build/parse CSS color strings and test fixtures
    // intentionally include hex colors.
    if (rel === "charts/scene/color.ts") return false;
    if (rel === "charts/scene/demo.ts") return false;
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
  // Avoid false positives like `rgb(...)` in comments by requiring either:
  // - a digit, or
  // - a decimal literal that starts with `.`, e.g. `.5`
  const rgbColor = /\brgb(a)?\s*\(\s*(?:\d|\.\d)/gi;

  // Named colors (e.g. `crimson`) bypass theme tokens. Allow only keyword-like values that
  // don't encode an actual palette color.
  const namedColorCss = /\bcolor\s*:\s*([a-zA-Z]+)\b(?!\s*\()/g;
  const namedColorJs = /\bcolor\s*:\s*["']([a-zA-Z]+)["']/g;
  const allowedNamedColors = new Set([
    "transparent",
    "inherit",
    "currentcolor",
    "initial",
    "unset",
    "revert",
  ]);

  /** @type {{ file: string, match: string }[]} */
  const violations = [];

  for (const file of files) {
    const content = fs.readFileSync(file, "utf8");
    const hex = content.match(hexColor);
    const rgb = content.match(rgbColor);
    if (hex) violations.push({ file, match: hex[0] });
    if (rgb) violations.push({ file, match: "rgb(...)" });

    // Flag named color literals (e.g. `color: "crimson"` / `color: crimson`) while
    // allowing keyword-like values such as `transparent`.
    if (file.endsWith(".css")) {
      for (const match of content.matchAll(namedColorCss)) {
        const value = match[1]?.toLowerCase();
        if (!value || allowedNamedColors.has(value)) continue;
        violations.push({ file, match: `color: ${match[1]}` });
        break;
      }
    } else {
      for (const match of content.matchAll(namedColorJs)) {
        const value = match[1]?.toLowerCase();
        if (!value || allowedNamedColors.has(value)) continue;
        violations.push({ file, match: `color: "${match[1]}"` });
        break;
      }
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded colors in core UI files:\\n${violations
      .map((v) => `- ${path.relative(srcRoot, v.file)}: ${v.match}`)
      .join("\\n")}`
  );
});
