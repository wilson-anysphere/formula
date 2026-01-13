import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Escape a string for safe interpolation into a RegExp.
 * @param {string} input
 */
function escapeRegExp(input) {
  return input.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

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

  // CSS also supports named colors (`crimson`, `red`, etc). These are disallowed in core UI;
  // use tokens instead (e.g. `var(--error)`), except for a few safe keywords.
  //
  // This list is intentionally *not* exhaustive; it's meant to catch high-signal offenders
  // while avoiding matching unrelated strings (e.g. "Red" as a UI label).
  const allowedColorKeywords = new Set(["transparent", "currentcolor", "inherit"]);
  const disallowedNamedColors = [
    "crimson",
    "red",
    "blue",
    "green",
    "orange",
    "yellow",
    "purple",
    "pink",
    "brown",
    "black",
    "white",
    "gray",
    "grey",
    "cyan",
    "magenta",
    "lime",
    "teal",
    "navy",
    "maroon",
    "olive",
    "silver",
    "gold",
  ];
  // Treat hyphenated identifiers as "words" so we don't match things like:
  // - `white-space`
  // - `var(--sheet-tab-red)`
  const namedColorToken = `(?<![\\w-])(?<color>${disallowedNamedColors.map(escapeRegExp).join("|")})(?![\\w-])`;
  const cssNamedColor = new RegExp(`:[^;{}]*${namedColorToken}`, "gi");
  const jsStyleColor = new RegExp(
    // style objects + style literals (e.g. `style={{ color: "red" }}`)
    String.raw`\b(?:accentColor|background|backgroundColor|border|borderColor|borderBottom|borderBottomColor|borderLeft|borderLeftColor|borderRight|borderRightColor|borderTop|borderTopColor|boxShadow|caretColor|color|fill|outline|outlineColor|stroke|textShadow)\b\s*:\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\1`,
    "gi",
  );
  const domStyleColor = new RegExp(
    // DOM style assignments (e.g. `el.style.color = "red"`)
    String.raw`\.style\.(?:accentColor|background|backgroundColor|borderColor|borderBottomColor|borderLeftColor|borderRightColor|borderTopColor|caretColor|color|fill|outlineColor|stroke)\s*=\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\1`,
    "gi",
  );
  const jsxAttributeColor = new RegExp(
    // JSX/SVG attrs (e.g. `<path fill="red" />`)
    String.raw`\b(?:accentColor|backgroundColor|borderColor|caretColor|color|fill|stroke|stopColor|floodColor|lightingColor)\b\s*=\s*(["'])[^"']*${namedColorToken}[^"']*\1`,
    "gi",
  );

  /** @type {{ file: string, match: string }[]} */
  const violations = [];

  for (const file of files) {
    const content = fs.readFileSync(file, "utf8");
    const hex = content.match(hexColor);
    const rgb = content.match(rgbColor);
    const ext = path.extname(file);
    /** @type {string | null} */
    let named = null;
    if (ext === ".css") {
      const match = cssNamedColor.exec(content);
      cssNamedColor.lastIndex = 0;
      named = match?.groups?.color ?? null;
    } else {
      const match =
        jsStyleColor.exec(content) ??
        domStyleColor.exec(content) ??
        jsxAttributeColor.exec(content);
      jsStyleColor.lastIndex = 0;
      domStyleColor.lastIndex = 0;
      jsxAttributeColor.lastIndex = 0;
      named = match?.groups?.color ?? null;
    }
    if (named && !allowedColorKeywords.has(named.toLowerCase())) {
      violations.push({ file, match: named });
    }
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
