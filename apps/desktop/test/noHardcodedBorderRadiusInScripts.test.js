import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";
import { stripComments } from "./sourceTextUtils.js";

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
  /** @type {Set<string>} */
  const borderRadiusCssVarRefs = new Set();
  const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;
  const unitRegex = /([+-]?(?:\d+(?:\.\d+)?|\.\d+))(px|%|rem|em|vh|vw|vmin|vmax|cm|mm|in|pt|pc|ch|ex)(?![A-Za-z0-9_])/gi;

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const stripped = stripComments(source);

    /** @type {{ re: RegExp, kind: string }[]} */
    const patterns = [
      // Style strings (e.g. `style: "border-radius: 4px;"`, `border-radius: calc(4px)`)
      { re: /\bborder-radius\s*:\s*(?<value>[^;"'`]*)/gi, kind: "border-radius" },
      // Longhand border radii in style strings (e.g. `border-top-left-radius: 4px`)
      {
        re: /\bborder-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\s*:\s*(?<value>[^;"'`]*)/gi,
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
      { re: /\.style\.borderRadius\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi, kind: "style.borderRadius-number" },
      // DOM style assignment via bracket notation (e.g. `el.style["borderRadius"] = 4`)
      {
        re: /\.style\s*\[\s*(?:["'`])borderRadius(?:["'`])\s*]\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style[borderRadius]-number",
      },
      // DOM style assignment for longhand border radii (numeric => px).
      {
        re: /\.style\.border(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style.border*Radius-number",
      },
      // DOM style assignment for longhand border radii via bracket notation (numeric => px).
      {
        re: /\.style\s*\[\s*(?:["'`])border(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius(?:["'`])\s*]\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style[border*Radius]-number",
      },
      // DOM style assignment (e.g. `el.style.borderRadius = "4px"`, `"calc(4px)"`)
      { re: /\.style\.borderRadius\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "style.borderRadius" },
      // DOM style assignment via bracket notation (e.g. `el.style["borderRadius"] = "4px"`, `"calc(4px)"`)
      {
        re: /\.style\s*\[\s*(?:["'`])borderRadius(?:["'`])\s*]\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi,
        kind: "style[borderRadius]",
      },
      // DOM style assignment for longhand border radii (string => px).
      {
        re: /\.style\.border(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi,
        kind: "style.border*Radius",
      },
      // DOM style assignment for longhand border radii via bracket notation (string => px).
      {
        re: /\.style\s*\[\s*(?:["'`])border(?:TopLeft|TopRight|BottomLeft|BottomRight|StartStart|StartEnd|EndStart|EndEnd)Radius(?:["'`])\s*]\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi,
        kind: "style[border*Radius]",
      },
      // setProperty("border-radius", 4)
      {
        re: /\.style\.setProperty\(\s*(["'])border-radius\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty-number",
      },
      // setProperty via bracket notation (e.g. `el.style["setProperty"]("border-radius", 4)`)
      {
        re: /\.style\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\(\s*(["'])border-radius\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty[border-radius]-number",
      },
      // setProperty("border-top-left-radius", 4)
      {
        re: /\.style\.setProperty\(\s*(["'])border-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty-border-*-radius-number",
      },
      // setProperty via bracket notation for longhand border radii (numeric) (e.g. `el.style["setProperty"]("border-top-left-radius", 4)`)
      {
        re: /\.style\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\(\s*(["'])border-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty[border-*-radius]-number",
      },
      // setProperty("border-radius", "4px") / setProperty(..., "calc(4px)")
      {
        re: /\.style\.setProperty\(\s*(["'])border-radius\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "setProperty",
      },
      // setProperty via bracket notation (e.g. `el.style["setProperty"]("border-radius", "4px")`)
      {
        re: /\.style\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\(\s*(["'])border-radius\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "setProperty[border-radius]",
      },
      // setProperty("border-top-left-radius", "4px")
      {
        re: /\.style\.setProperty\(\s*(["'])border-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "setProperty-border-*-radius",
      },
      // setProperty via bracket notation for longhand border radii (e.g. `el.style["setProperty"]("border-top-left-radius", "4px")`)
      {
        re: /\.style\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\(\s*(["'])border-(?:top|bottom|start|end)-(?:left|right|start|end)-radius\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "setProperty[border-*-radius]",
      },
    ];

    for (const { re } of patterns) {
      let match;
      while ((match = re.exec(stripped))) {
        const valueString = match.groups?.value;
        if (typeof valueString === "string") {
          // Capture CSS vars referenced by border-radius declarations so hardcoded units cannot be
          // hidden behind custom properties (possibly in a different stylesheet).
          let varMatch;
          while ((varMatch = cssVarRef.exec(valueString))) {
            borderRadiusCssVarRefs.add(varMatch[1]);
          }
          cssVarRef.lastIndex = 0;

          // Scan the matched value for any hardcoded length units (e.g. `calc(4px)` or `var(--radius, 4px)`).
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
          unitRegex.lastIndex = 0;
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

  // Also ensure that any CSS variables used by border-radius assignments in scripts do not themselves
  // hide hardcoded unit values (except canonical `--radius*` tokens).
  const nonTokenVarRefs = [...borderRadiusCssVarRefs].filter((ref) => !ref.startsWith("--radius"));
  if (nonTokenVarRefs.length > 0) {
    /**
     * @param {string} dirPath
     * @returns {string[]}
     */
    function walkCssFiles(dirPath) {
      /** @type {string[]} */
      const files = [];
      for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
        const fullPath = path.join(dirPath, entry.name);
        if (entry.isDirectory()) {
          files.push(...walkCssFiles(fullPath));
          continue;
        }
        if (!entry.isFile()) continue;
        if (!entry.name.endsWith(".css")) continue;
        files.push(fullPath);
      }
      return files;
    }

    const cssFiles = walkCssFiles(srcRoot).filter((file) => {
      const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
      // Demo/sandbox assets are not part of the shipped UI bundle.
      if (rel.startsWith("grid/presence-renderer/")) return false;
      if (rel.includes("/demo/")) return false;
      if (rel.includes("/__tests__/")) return false;
      return true;
    });

    /** @type {Map<string, string>} */
    const strippedByFile = new Map();
    /** @type {Map<string, Array<{ file: string, value: string, valueStart: number }>>} */
    const customPropDecls = new Map();
    const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;

    for (const file of cssFiles) {
      const css = fs.readFileSync(file, "utf8");
      const strippedCss = stripCssNonSemanticText(css);
      strippedByFile.set(file, strippedCss);

      cssDeclaration.lastIndex = 0;
      let decl;
      while ((decl = cssDeclaration.exec(strippedCss))) {
        const prop = decl?.groups?.prop ?? "";
        if (!prop.startsWith("--")) continue;

        const value = decl?.groups?.value ?? "";
        const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

        let entries = customPropDecls.get(prop);
        if (!entries) {
          entries = [];
          customPropDecls.set(prop, entries);
        }
        entries.push({ file, value, valueStart });
      }
      cssDeclaration.lastIndex = 0;
    }

    /** @type {Set<string>} */
    const expandedVarRefs = new Set(nonTokenVarRefs);
    const queue = [...nonTokenVarRefs];
    while (queue.length > 0) {
      const varName = queue.pop();
      if (varName.startsWith("--radius")) continue;
      const declsForVar = customPropDecls.get(varName);
      if (!declsForVar) continue;

      for (const { value } of declsForVar) {
        let varMatch;
        while ((varMatch = cssVarRef.exec(value))) {
          const ref = varMatch[1];
          if (expandedVarRefs.has(ref)) continue;
          expandedVarRefs.add(ref);
          queue.push(ref);
        }
        cssVarRef.lastIndex = 0;
      }
    }

    for (const [prop, declsForVar] of customPropDecls) {
      if (!expandedVarRefs.has(prop)) continue;
      if (prop.startsWith("--radius")) continue;

      for (const { file, value, valueStart } of declsForVar) {
        const strippedCss = strippedByFile.get(file) ?? "";
        let unitMatch;
        while ((unitMatch = unitRegex.exec(value))) {
          const numeric = unitMatch[1];
          const unit = unitMatch[2] ?? "";
          const n = Number(numeric);
          if (!Number.isFinite(n)) continue;
          if (n === 0) continue;

          const absIndex = valueStart + unitMatch.index;
          const line = getLineNumber(strippedCss, absIndex);
          const rawUnit = unitMatch[0] ?? `${numeric}${unit}`;
          violations.push(
            `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${prop}: ${value.trim()} (found ${rawUnit})`,
          );
        }

        unitRegex.lastIndex = 0;
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
