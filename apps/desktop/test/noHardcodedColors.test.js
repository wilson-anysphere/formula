import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";
import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

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

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
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
    // Table formatting presets intentionally embed OOXML/Excel ARGB strings (cell formatting).
    if (rel === "formatting/formatAsTablePresets.ts") return false;
    // Charts scene graph utilities build/parse CSS color strings and test fixtures
    // intentionally include hex colors.
    if (rel === "charts/scene/color.ts") return false;
    if (rel === "charts/scene/demo.ts") return false;
    if (rel.includes("/demo/")) return false;
    // DrawingML parsers may emit/document literal hex colors derived from document
    // payloads (not theme/UI tokens).
    if (rel === "drawings/shapeRenderer.ts") return false;
    if (rel.includes("/__tests__/")) return false;
    // Vitest entrypoints live under `src/` with a `.vitest.*` suffix; treat them like other tests.
    // These files are not part of the shipped UI bundle and frequently contain hardcoded color
    // literals in fixtures/assertions.
    if (/\.(test|spec|vitest)\.[jt]sx?$/.test(rel)) return false;
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
  // Same logic for hsl()/hsla(). Require a numeric channel so we don't match parsing helpers like
  // `hsl()` in comments, regex literals, or `hsl(var(--foo))`. Include an optional sign so
  // hardcoded hues like `hsl(-30deg ...)` can't slip through.
  const hslColor = /\bhsl(a)?\s*\(\s*(?:[+-]?(?:\d|\.\d))/gi;
  // Catch other CSS color functions (less common but still hardcoded colors). Require a numeric
  // channel for the same "high signal" reason as rgb/hsl.
  const modernColorFn = /\b(?<fn>hwb|lab|lch|oklab|oklch)\s*\(\s*(?:[+-]?(?:\d|\.\d))/gi;

  // CSS also supports named colors (`crimson`, `red`, etc). These are disallowed in core UI;
  // use tokens instead (e.g. `var(--error)`), except for a few safe keywords.
  //
  // This list is intentionally *not* exhaustive; it's meant to catch high-signal offenders
  // while avoiding matching unrelated strings (e.g. "Red" as a UI label).
  const allowedColorKeywords = new Set(["transparent", "currentcolor", "inherit", "initial", "unset", "revert"]);
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
  // Also treat `/` and `.` as "word-ish" boundaries so we don't match file paths like `icons/red.png`.
  const namedColorToken = `(?<![\\w\\-/.])(?<color>${disallowedNamedColors.map(escapeRegExp).join("|")})(?![\\w\\-/.])`;
  const namedColorTokenRe = new RegExp(namedColorToken, "gi");
  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
  const cssColorProps = new Set([
    "accent-color",
    "background",
    "background-color",
    "background-image",
    "border",
    "border-color",
    "border-top",
    "border-top-color",
    "border-right",
    "border-right-color",
    "border-bottom",
    "border-bottom-color",
    "border-left",
    "border-left-color",
    "box-shadow",
    "caret-color",
    "color",
    "fill",
    "outline",
    "outline-color",
    "stroke",
    "text-shadow",
  ]);
  const jsStyleColor = new RegExp(
    // style objects + style literals (e.g. `style={{ color: "red" }}`)
    String.raw`\b(?:accentColor|background|backgroundColor|backgroundImage|border|borderColor|borderBottom|borderBottomColor|borderLeft|borderLeftColor|borderRight|borderRightColor|borderTop|borderTopColor|boxShadow|caretColor|color|fill|fillStyle|outline|outlineColor|stroke|strokeStyle|textShadow)\b\s*:\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\1`,
    "gi",
  );
  const domStyleColor = new RegExp(
    // DOM style assignments (e.g. `el.style.color = "red"`)
    String.raw`\.style\.(?:accentColor|background|backgroundColor|borderColor|borderBottomColor|borderLeftColor|borderRightColor|borderTopColor|caretColor|color|fill|outlineColor|stroke)\s*=\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\1`,
    "gi",
  );
  const canvasStyleColor = new RegExp(
    // CanvasRenderingContext2D assignments (e.g. `ctx.fillStyle = "red"`)
    String.raw`\b(?:fillStyle|strokeStyle)\b\s*=\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\1`,
    "gi",
  );
  const setPropertyStyleColor = new RegExp(
    // DOM style setProperty assignments (e.g. `el.style.setProperty("color", "red")` or `setProperty("--foo", "red")`)
    String.raw`\.style\.setProperty\(\s*(["'\`])(?<prop>[-\w]+)\1\s*,\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\3`,
    "gi",
  );
  const setAttributeColor = new RegExp(
    // SVG/DOM attribute assignments (e.g. `el.setAttribute("fill", "red")`)
    String.raw`\bsetAttribute\(\s*(["'])(?:fill|stroke|color|stop-color|flood-color|lighting-color)\1\s*,\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\2`,
    "gi",
  );
  const jsxAttributeColor = new RegExp(
    // JSX/SVG attrs (e.g. `<path fill="red" />`)
    String.raw`\b(?:accentColor|backgroundColor|borderColor|caretColor|color|fill|stroke|stopColor|floodColor|lightingColor)\b\s*=\s*(["'])[^"']*${namedColorToken}[^"']*\1`,
    "gi",
  );
  const jsxAttributeColorExpr = new RegExp(
    // JSX/SVG attrs with string literal expressions (e.g. `<path fill={"red"} />`)
    String.raw`\b(?:accentColor|backgroundColor|borderColor|caretColor|color|fill|stroke|stopColor|floodColor|lightingColor)\b\s*=\s*\{\s*(["'\`])[^"'\`]*${namedColorToken}[^"'\`]*\1\s*\}`,
    "gi",
  );

  /** @type {string[]} */
  const violations = [];

  for (const file of files) {
    const content = fs.readFileSync(file, "utf8");
    const ext = path.extname(file);
    const stripped = ext === ".css" ? stripCssNonSemanticText(content) : stripComments(content);
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");

    hexColor.lastIndex = 0;
    rgbColor.lastIndex = 0;
    hslColor.lastIndex = 0;
    modernColorFn.lastIndex = 0;

    const hex = hexColor.exec(stripped);
    const rgb = rgbColor.exec(stripped);
    const hsl = hslColor.exec(stripped);
    const modern = modernColorFn.exec(stripped);

    hexColor.lastIndex = 0;
    rgbColor.lastIndex = 0;
    hslColor.lastIndex = 0;
    modernColorFn.lastIndex = 0;

    /** @type {string | null} */
    let named = null;
    /** @type {number | null} */
    let namedIndex = null;
    /** @type {string} */
    let namedContext = "";
    if (ext === ".css") {
      for (const decl of stripped.matchAll(cssDeclaration)) {
        const prop = decl?.groups?.prop?.toLowerCase() ?? "";
        const value = decl?.groups?.value ?? "";
        if (!prop) continue;
        const isCustomProp = prop.startsWith("--");
        if (!isCustomProp && !cssColorProps.has(prop)) continue;
        const match = namedColorTokenRe.exec(value);
        namedColorTokenRe.lastIndex = 0;
        if (match?.groups?.color) {
          named = match.groups.color;
          const matchStart = decl.index ?? 0;
          const valueStart = matchStart + (decl[0]?.length ?? 0) - value.length;
          namedIndex = valueStart + (match.index ?? 0);
          namedContext = prop;
          break;
        }
      }
    } else {
      const match =
        jsStyleColor.exec(stripped) ??
        domStyleColor.exec(stripped) ??
        canvasStyleColor.exec(stripped) ??
        setPropertyStyleColor.exec(stripped) ??
        setAttributeColor.exec(stripped) ??
        jsxAttributeColor.exec(stripped) ??
        jsxAttributeColorExpr.exec(stripped);
      jsStyleColor.lastIndex = 0;
      domStyleColor.lastIndex = 0;
      canvasStyleColor.lastIndex = 0;
      setPropertyStyleColor.lastIndex = 0;
      setAttributeColor.lastIndex = 0;
      jsxAttributeColor.lastIndex = 0;
      jsxAttributeColorExpr.lastIndex = 0;
      named = match?.groups?.color ?? null;
      if (named) {
        namedIndex = (match?.index ?? 0) + (match?.[0]?.indexOf(named) ?? 0);
        const prop = match?.groups?.prop;
        namedContext = prop ? `setProperty(${prop})` : "named-color";
      }
    }
    if (named && !allowedColorKeywords.has(named.toLowerCase())) {
      const line = getLineNumber(stripped, namedIndex ?? 0);
      violations.push(`${rel}:L${line}: ${namedContext}: ${named}`);
    }
    if (hex) {
      const line = getLineNumber(stripped, hex.index ?? 0);
      violations.push(`${rel}:L${line}: ${hex[0]}`);
    }
    if (rgb) {
      const line = getLineNumber(stripped, rgb.index ?? 0);
      violations.push(`${rel}:L${line}: rgb(...)`);
    }
    if (hsl) {
      const line = getLineNumber(stripped, hsl.index ?? 0);
      violations.push(`${rel}:L${line}: hsl(...)`);
    }
    if (modern) {
      const line = getLineNumber(stripped, modern.index ?? 0);
      violations.push(`${rel}:L${line}: ${modern.groups?.fn ?? "color"}(...)`);
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded colors in core UI files:\\n${violations.map((v) => `- ${v}`).join("\\n")}`,
  );
});
