import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const srcRoot = path.join(desktopRoot, "src");

/**
 * @param {string} dirPath
 * @returns {string[]}
 */
function walkScriptFiles(dirPath) {
  /** @type {string[]} */
  const files = [];
  for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      files.push(...walkScriptFiles(fullPath));
      continue;
    }
    if (!entry.isFile()) continue;
    if (!/\.[jt]sx?$/.test(entry.name)) continue;
    files.push(fullPath);
  }
  return files;
}

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

/**
 * Strip JS/TS line + block comments while preserving string literals.
 *
 * This matches the approach used by noHardcodedColors.test.js so guardrails stay
 * high-signal without attempting to fully parse JavaScript.
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

    // String literals: preserve as-is so we can scan inline style strings.
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

test("desktop UI scripts should not hardcode border-radius values in inline styles", () => {
  const files = walkScriptFiles(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    // Demo/sandbox assets are not part of the shipped UI bundle.
    if (rel.startsWith("grid/presence-renderer/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    if (/\.(test|spec|vitest)\.[jt]sx?$/.test(rel)) return false;
    return true;
  });

  /** @type {string[]} */
  const violations = [];

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const stripped = stripJsComments(source);

    /** @type {{ re: RegExp, kind: string }[]} */
    const patterns = [
      // Style strings (e.g. `style: "border-radius: 4px;"`, `border-radius: calc(4px)`)
      { re: /\bborder-radius\s*:\s*(?<value>[^;"'`}]*)/gi, kind: "border-radius" },
      // Longhand border radii in style strings (e.g. `border-top-left-radius: 4px`)
      {
        re: /\bborder-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\s*:\s*(?<value>[^;"'`}]*)/gi,
        kind: "border-*-radius",
      },
      // React style objects (e.g. `{ borderRadius: 4 }`) interpret numeric values as px.
      { re: /\bborderRadius\s*:\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi, kind: "borderRadius-number" },
      // Longhand border radii in React style objects (numeric => px).
      {
        re: /\bborder(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius\s*:\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "border*Radius-number",
      },
      // React/DOM style objects (e.g. `{ borderRadius: "4px" }`, `"calc(4px)"`)
      { re: /\bborderRadius\s*:\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "borderRadius" },
      // Longhand border radii in React style objects (string => px).
      {
        re: /\bborder(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius\s*:\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi,
        kind: "border*Radius",
      },
      // DOM style assignment (e.g. `el.style.borderRadius = 4`)
      { re: /\.style\.borderRadius\s*=\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi, kind: "style.borderRadius-number" },
      // DOM style assignment for longhand border radii (numeric => px).
      {
        re: /\.style\.border(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius\s*=\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style.border*Radius-number",
      },
      // DOM style assignment (e.g. `el.style.borderRadius = "4px"`, `"calc(4px)"`)
      { re: /\.style\.borderRadius\s*=\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "style.borderRadius" },
      // DOM style assignment for longhand border radii (string => px).
      {
        re: /\.style\.border(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius\s*=\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi,
        kind: "style.border*Radius",
      },
      // setProperty("border-radius", 4)
      {
        re: /\.style\.setProperty\(\s*(["'])border-radius\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty-number",
      },
      // setProperty("border-top-left-radius", 4)
      {
        re: /\.style\.setProperty\(\s*(["'])border-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty-border-*-radius-number",
      },
      // setProperty("border-radius", "4px") / setProperty(..., "calc(4px)")
      {
        re: /\.style\.setProperty\(\s*(["'])border-radius\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "setProperty",
      },
      // setProperty("border-top-left-radius", "4px")
      {
        re: /\.style\.setProperty\(\s*(["'])border-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "setProperty-border-*-radius",
      },
    ];

    for (const { re } of patterns) {
      let match;
      while ((match = re.exec(stripped))) {
        const valueString = match.groups?.value;
        if (typeof valueString === "string") {
          // Scan the matched value for any hardcoded length units (e.g. `calc(4px)` or `var(--radius, 4px)`).
          const unitRegex =
            /([+-]?(?:\d+(?:\.\d+)?|\.\d+))(px|%|rem|em|vh|vw|vmin|vmax|cm|mm|in|pt|pc|ch|ex)(?![A-Za-z0-9_])/gi;
          const valueStart = match[0].indexOf(valueString);
          let unitMatch;
          while ((unitMatch = unitRegex.exec(valueString))) {
            const numeric = unitMatch[1];
            const unit = unitMatch[2] ?? "";
            const n = Number(numeric);
            if (!Number.isFinite(n)) continue;
            if (n === 0) continue;

            const absIndex = match.index + Math.max(0, valueStart) + unitMatch.index;
            const line = getLineNumber(stripped, absIndex);
            violations.push(
              `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: border-radius: ${numeric}${unit}`,
            );
          }
          continue;
        }

        const numeric = match.groups?.num;
        if (!numeric) continue;
        const px = Number(numeric);
        if (px === 0) continue;

        // Find the absolute index of the numeric capture for stable line numbers.
        const needle = String(numeric);
        const relative = match[0].indexOf(needle);
        const absIndex = match.index + (relative >= 0 ? relative : 0);
        const line = getLineNumber(stripped, absIndex);
        violations.push(`${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: border-radius: ${numeric}px`);
      }
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded border-radius values in desktop UI scripts. Use radius tokens (var(--radius*)), except for 0:\n${violations
      .map((v) => `- ${v}`)
      .join("\n")}`,
  );
});
