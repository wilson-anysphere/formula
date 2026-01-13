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

/**
 * Remove CSS string literals, comments, and url(...) bodies so we don't match named colors
 * inside unrelated text (e.g. asset paths like `url("/icons/red.png")`).
 *
 * This is not a full CSS parser; it's just enough to keep the `noHardcodedColors` guardrail
 * high-signal while avoiding common false positives.
 *
 * @param {string} css
 */
function stripCssNonSemanticText(css) {
  let out = String(css);
  // Block comments.
  out = out.replace(/\/\*[\s\S]*?\*\//g, " ");
  // Quoted strings (handles escapes).
  out = out.replace(/"(?:\\.|[^"\\])*"/g, '""');
  out = out.replace(/'(?:\\.|[^'\\])*'/g, "''");

  // Strip url(...) bodies, handling nested parens in a minimal way and respecting quotes.
  let idx = 0;
  let result = "";
  while (idx < out.length) {
    const m = /\burl\s*\(/gi.exec(out.slice(idx));
    if (!m) {
      result += out.slice(idx);
      break;
    }
    const start = idx + (m.index ?? 0);
    result += out.slice(idx, start);

    // Find the opening '(' we matched.
    const openParen = out.indexOf("(", start);
    if (openParen === -1) {
      // Shouldn't happen, but fall back to copying the rest.
      result += out.slice(start);
      break;
    }

    let i = openParen + 1;
    let depth = 1;
    /** @type {string | null} */
    let quote = null;
    while (i < out.length && depth > 0) {
      const ch = out[i];
      if (quote) {
        if (ch === "\\") {
          i += 2;
          continue;
        }
        if (ch === quote) {
          quote = null;
        }
        i += 1;
        continue;
      }
      if (ch === '"' || ch === "'") {
        quote = ch;
        i += 1;
        continue;
      }
      if (ch === "(") depth += 1;
      else if (ch === ")") depth -= 1;
      i += 1;
    }

    // Replace the entire url(...) with a stub so declaration parsing stays roughly intact.
    result += "url()";
    idx = i;
  }

  return result;
}

/**
 * Strip JS/TS line + block comments while preserving string literals.
 *
 * This keeps the named-color scan high-signal (don't flag `red` mentioned in docs/JSDoc)
 * without attempting to fully parse JavaScript.
 *
 * @param {string} input
 */
function stripJsComments(input) {
  const text = String(input);
  let out = "";
  /** @type {"code" | "single" | "double" | "template" | "lineComment" | "blockComment"} */
  let state = "code";

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = i + 1 < text.length ? text[i + 1] : "";

    if (state === "code") {
      if (ch === "'" || ch === '"' || ch === "`") {
        state = ch === "'" ? "single" : ch === '"' ? "double" : "template";
        out += ch;
        continue;
      }

      if (ch === "/" && next === "/") {
        state = "lineComment";
        out += "  ";
        i += 1;
        continue;
      }

      if (ch === "/" && next === "*") {
        state = "blockComment";
        out += "  ";
        i += 1;
        continue;
      }

      out += ch;
      continue;
    }

    if (state === "lineComment") {
      if (ch === "\n") {
        state = "code";
        out += "\n";
      } else {
        out += " ";
      }
      continue;
    }

    if (state === "blockComment") {
      if (ch === "*" && next === "/") {
        state = "code";
        out += "  ";
        i += 1;
        continue;
      }
      out += ch === "\n" ? "\n" : " ";
      continue;
    }

    // String literals: preserve as-is so style scans can still match string values.
    out += ch;
    if (ch === "\\") {
      if (next) {
        out += next;
        i += 1;
      }
      continue;
    }

    if (state === "single" && ch === "'") {
      state = "code";
    } else if (state === "double" && ch === '"') {
      state = "code";
    } else if (state === "template" && ch === "`") {
      state = "code";
    }
  }

  return out;
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
    // Ribbon command router maps ribbon ids to cell-formatting color presets (data-driven),
    // which intentionally include ARGB hex strings.
    if (rel === "ribbon/ribbonCommandRouter.ts") return false;
    if (rel.includes("/demo/")) return false;
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
      const stripped = stripCssNonSemanticText(content);
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
          break;
        }
      }
    } else {
      const stripped = stripJsComments(content);
      const match =
        jsStyleColor.exec(stripped) ??
        domStyleColor.exec(stripped) ??
        canvasStyleColor.exec(stripped) ??
        setAttributeColor.exec(stripped) ??
        jsxAttributeColor.exec(stripped);
      jsStyleColor.lastIndex = 0;
      domStyleColor.lastIndex = 0;
      canvasStyleColor.lastIndex = 0;
      setAttributeColor.lastIndex = 0;
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
