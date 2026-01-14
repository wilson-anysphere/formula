import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments, stripCssComments } from "./sourceTextUtils.js";

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

test("desktop UI scripts should not hardcode monospace font stacks in inline styles (use --font-mono)", () => {
  const files = walkScriptFiles(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    // Demo/sandbox assets are not part of the shipped UI bundle.
    if (rel.startsWith("grid/presence-renderer/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    if (/\.(test|spec|vitest)\.[jt]sx?$/.test(rel)) return false;
    return true;
  });

  // Keep in sync with `--font-mono` in `src/styles/tokens.css`.
  const forbiddenFontStackToken = /\b(?:ui-monospace|SFMono(?:-Regular)?|SF\s*Mono|Menlo|Consolas|monospace)\b/gi;
  const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;
  const allowedFontVars = new Set(["--font-mono", "--font-sans"]);

  /** @type {string[]} */
  const violations = [];
  /** @type {Set<string>} */
  const fontVarRefs = new Set();

  /** @type {{ re: RegExp, kind: string }[]} */
  const patterns = [
    // CSS declarations embedded in style strings (e.g. `style: "font-family: ui-monospace"`).
    { re: /\bfont-family\s*:\s*(?<value>[^;"'`]*)/gi, kind: "font-family" },
    // CSS declarations embedded in style strings (e.g. `style: "font: 12px ui-monospace"`).
    { re: /\bfont\s*:\s*(?<value>[^;"'`]*)/gi, kind: "font" },
    // React style objects (e.g. `{ fontFamily: "ui-monospace" }`).
    { re: /\bfontFamily\s*:\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "fontFamily" },
    // React style objects (e.g. `{ font: "12px ui-monospace" }`).
    { re: /\bfont\s*:\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "font" },
    // DOM style assignment (e.g. `el.style.fontFamily = "ui-monospace"`).
    { re: /\.style\.fontFamily\s*=\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "style.fontFamily" },
    // DOM style assignment (e.g. `el.style.font = "12px ui-monospace"`).
    { re: /\.style\.font\s*=\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi, kind: "style.font" },
    // setProperty("font-family", "ui-monospace")
    {
      re: /\.style\.setProperty\(\s*(["'])font-family\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi,
      kind: "setProperty(font-family)",
    },
    // setProperty("font", "12px ui-monospace")
    { re: /\.style\.setProperty\(\s*(["'])font\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi, kind: "setProperty(font)" },
  ];

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const stripped = stripComments(source);
    const rel = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

    for (const { re, kind } of patterns) {
      re.lastIndex = 0;
      let match;
      while ((match = re.exec(stripped))) {
        const value = match.groups?.value ?? "";
        if (!value) continue;

        // Collect any CSS variables referenced by font-family assignments so monospace stacks cannot
        // be hidden behind a custom property defined elsewhere.
        let varMatch;
        while ((varMatch = cssVarRef.exec(value))) {
          fontVarRefs.add(varMatch[1]);
        }
        cssVarRef.lastIndex = 0;

        forbiddenFontStackToken.lastIndex = 0;
        let tokenMatch;
        while ((tokenMatch = forbiddenFontStackToken.exec(value))) {
          const matchStart = match.index ?? 0;
          const matchText = match[0] ?? "";
          const token = tokenMatch[0] ?? "";
          const valueOffset = matchText.indexOf(value);
          const absIndex = matchStart + (valueOffset >= 0 ? valueOffset : 0) + (tokenMatch.index ?? 0);
          const line = getLineNumber(stripped, absIndex);
          violations.push(`${rel}:L${line}: ${kind}: ${token}`);
        }
      }
    }
  }

  // Also ensure that any CSS variables used by script-set font-family declarations do not themselves
  // hide hardcoded monospace stacks (except canonical `--font-mono`).
  const nonTokenVarRefs = [...fontVarRefs].filter((ref) => !allowedFontVars.has(ref));
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
      const raw = fs.readFileSync(file, "utf8");
      const strippedCss = stripCssComments(raw);
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
      if (!varName) continue;
      if (allowedFontVars.has(varName)) continue;
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
      if (allowedFontVars.has(prop)) continue;

      for (const { file, value, valueStart } of declsForVar) {
        const strippedCss = strippedByFile.get(file) ?? "";
        const rel = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

        forbiddenFontStackToken.lastIndex = 0;
        let tokenMatch;
        while ((tokenMatch = forbiddenFontStackToken.exec(value))) {
          const absIndex = valueStart + (tokenMatch.index ?? 0);
          const line = getLineNumber(strippedCss, absIndex);
          const token = tokenMatch[0] ?? "";
          violations.push(`${rel}:L${line}: ${prop}: ${value.trim()} (found ${token})`);
        }
      }
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded monospace font stacks in desktop UI scripts. Use var(--font-mono) from src/styles/tokens.css:\n${violations
      .map((v) => `- ${v}`)
      .join("\n")}`,
  );
});
