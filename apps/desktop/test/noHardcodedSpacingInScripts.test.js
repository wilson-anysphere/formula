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

test("desktop UI scripts should not hardcode px values for padding/margin/gap in inline styles", () => {
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
  const spacingCssVarRefs = new Set();

  const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;
  const pxUnit = /([+-]?(?:\d+(?:\.\d+)?|\.\d+))px(?![A-Za-z0-9_])/gi;

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const stripped = stripComments(source);

    /** @type {{ re: RegExp, kind: string }[]} */
    const patterns = [
      // CSS declarations inside style strings (e.g. `style: "padding: 8px;"`).
      {
        re: /\b(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\s*:\s*(?<value>[^;"'`]*)/gi,
        kind: "css-declaration",
      },
      // DOM style assignment (e.g. `el.style.marginTop = "8px"`).
      {
        re: /\.\s*style\b\s*\.\s*(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "style.prop",
      },
      // DOM style assignment via bracket access to the `style` property (e.g. `el["style"].marginTop = "8px"`).
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*\.\s*(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "style['style'].prop",
      },
      // DOM style assignment via bracket notation (e.g. `el.style["marginTop"] = "8px"`).
      {
        re: /\.\s*style\b\s*\[\s*(?:["'`])(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:["'`])\s*]\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "style[prop]",
      },
      // DOM style assignment via bracket access to `style` + bracket notation (e.g. `el["style"]["marginTop"] = "8px"`).
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*\[\s*(?:["'`])(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:["'`])\s*]\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
        kind: "style['style'][prop]",
      },
      // DOM style assignment (numeric) (e.g. `el.style.marginTop = 8`).
      {
        re: /\.\s*style\b\s*\.\s*(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style.prop-number",
      },
      // DOM style assignment via bracket access to `style` (numeric) (e.g. `el["style"].marginTop = 8`).
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*\.\s*(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style['style'].prop-number",
      },
      // DOM style assignment via bracket notation (numeric) (e.g. `el.style["marginTop"] = 8`).
      {
        re: /\.\s*style\b\s*\[\s*(?:["'`])(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:["'`])\s*]\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style[prop]-number",
      },
      // DOM style assignment via bracket access to `style` + bracket notation (numeric) (e.g. `el["style"]["marginTop"] = 8`).
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*\[\s*(?:["'`])(?<prop>gap|rowGap|columnGap|padding(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?|margin(?:Top|Right|Bottom|Left|Inline|InlineStart|InlineEnd|Block|BlockStart|BlockEnd)?)\s*(?:["'`])\s*]\s*(?:=|\+=)\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "style['style'][prop]-number",
      },
      // setProperty("margin-top", "8px")
      {
        re: /\.\s*style\b\s*(?:\?\.|\.)\s*setProperty\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\3/gi,
        kind: "setProperty",
      },
      // setProperty via bracket access to `style` (e.g. `el["style"].setProperty("margin-top", "8px")`)
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*(?:\?\.|\.)\s*setProperty\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\3/gi,
        kind: "setProperty['style']",
      },
      // setProperty via bracket notation (e.g. `el.style["setProperty"]("margin-top", "8px")`)
      {
        re: /\.\s*style\b\s*(?:\?\.)?\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\3/gi,
        kind: "setProperty[prop]",
      },
      // setProperty via bracket access to `style` + bracket notation (e.g. `el["style"]["setProperty"]("margin-top", "8px")`)
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*(?:\?\.)?\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\3/gi,
        kind: "setProperty['style'][prop]",
      },
      // setProperty("margin-top", 8)
      {
        re: /\.\s*style\b\s*(?:\?\.|\.)\s*setProperty\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty-number",
      },
      // setProperty via bracket access to `style` (numeric) (e.g. `el["style"].setProperty("margin-top", 8)`)
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*(?:\?\.|\.)\s*setProperty\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty['style']-number",
      },
      // setProperty via bracket notation (numeric) (e.g. `el.style["setProperty"]("margin-top", 8)`)
      {
        re: /\.\s*style\b\s*(?:\?\.)?\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty[prop]-number",
      },
      // setProperty via bracket access to `style` + bracket notation (numeric) (e.g. `el["style"]["setProperty"]("margin-top", 8)`)
      {
        re: /\[\s*(?:["'`])style(?:["'`])\s*]\s*(?:\?\.)?\s*\[\s*(?:["'`])setProperty(?:["'`])\s*]\s*(?:\(\s*|\.\s*call\s*\(\s*[^,]+,\s*)(["'`])(?<prop>gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)\1\s*,\s*(?<num>[+-]?(?:\d+(?:\.\d+)?|\.\d+))\b/gi,
        kind: "setProperty['style'][prop]-number",
      },
    ];

    for (const { re, kind } of patterns) {
      let match;
      while ((match = re.exec(stripped))) {
        const valueString = match.groups?.value;
        const prop = match.groups?.prop ?? "";

        if (typeof valueString === "string") {
          // Capture CSS vars referenced by spacing declarations so hardcoded px values cannot be
          // hidden behind custom properties (possibly in a different stylesheet).
          let varMatch;
          while ((varMatch = cssVarRef.exec(valueString))) {
            spacingCssVarRefs.add(varMatch[1]);
          }
          cssVarRef.lastIndex = 0;

          const valueStart = match[0].indexOf(valueString);
          let unitMatch;
          while ((unitMatch = pxUnit.exec(valueString))) {
            const numeric = unitMatch[1] ?? "";
            const n = Number(numeric);
            if (!Number.isFinite(n)) continue;
            if (n === 0) continue;

            const absIndex = match.index + Math.max(0, valueStart) + (unitMatch.index ?? 0);
            const line = getLineNumber(stripped, absIndex);
            const rawUnit = unitMatch[0] ?? `${numeric}px`;
            violations.push(
              `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${kind} ${prop}: ${valueString.trim()} (found ${rawUnit})`,
            );
          }
          pxUnit.lastIndex = 0;
          continue;
        }

        const numeric = match.groups?.num;
        if (!numeric) continue;
        const n = Number(numeric);
        if (!Number.isFinite(n)) continue;
        if (n === 0) continue;

        const needle = String(numeric);
        const relative = match[0].indexOf(needle);
        const absIndex = match.index + (relative >= 0 ? relative : 0);
        const line = getLineNumber(stripped, absIndex);
        violations.push(
          `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${kind} ${prop}: ${numeric}px`,
        );
      }
    }
  }

  // Also ensure that any CSS variables used by script-set spacing declarations do not themselves
  // hide hardcoded px values (except canonical `--space-*` tokens).
  const nonTokenVarRefs = [...spacingCssVarRefs].filter((ref) => !ref.startsWith("--space-"));
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
      if (varName.startsWith("--space-")) continue;
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
      if (prop.startsWith("--space-")) continue;

      for (const { file, value, valueStart } of declsForVar) {
        const strippedCss = strippedByFile.get(file) ?? "";
        let unitMatch;
        while ((unitMatch = pxUnit.exec(value))) {
          const numeric = unitMatch[1] ?? "";
          const n = Number(numeric);
          if (!Number.isFinite(n)) continue;
          if (n === 0) continue;

          const absIndex = valueStart + (unitMatch.index ?? 0);
          const line = getLineNumber(strippedCss, absIndex);
          const rawUnit = unitMatch[0] ?? `${numeric}px`;
          violations.push(
            `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${prop}: ${value.trim()} (used by scripts; found ${rawUnit})`,
          );
        }
        pxUnit.lastIndex = 0;
      }
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded px values in desktop UI scripts for padding/margin/gap (use --space-* tokens; allow 0):\n${violations
      .map((v) => `- ${v}`)
      .join("\n")}`,
  );
});
