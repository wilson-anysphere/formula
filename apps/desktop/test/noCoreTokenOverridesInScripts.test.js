import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
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

test("desktop UI scripts should not override core design tokens", () => {
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

  // Only match direct `.style.setProperty("--token", ...)` / `.style.removeProperty("--token")`
  // calls so we don't trigger on other string usages like `getPropertyValue("--space-4")`.
  const tokenOp =
    /\.\s*style\b\s*(?:\?\.|\.)\s*(?<op>setProperty|removeProperty)\s*\(\s*(["'`])\s*(?<prop>--(?:space-\d+|radius[-\w]*|motion-(?:duration(?:-fast)?|ease)|font-(?:mono|sans)|bg-[\w-]+|text-[\w-]+|border[\w-]*|accent[\w-]*|selection-[\w-]+|titlebar-[\w-]+|sheet-tab-[\w-]+|chart-[\w-]+|tooltip-[\w-]+|cmdk-[\w-]+|shadow-[\w-]+|formula-[\w-]+|grid-header-[\w-]+|grid-line|panel-(?:bg|border|shadow)|dialog-(?:bg|border|shadow|backdrop)|link|error[\w-]*|warning[\w-]*|success[\w-]*))\s*\2/gi;
  const tokenOpBracket =
    /\.\s*style\b\s*(?:\?\.)?\s*\[\s*(?:["'`])(?<op>setProperty|removeProperty)(?:["'`])\s*]\s*\(\s*(["'`])\s*(?<prop>--(?:space-\d+|radius[-\w]*|motion-(?:duration(?:-fast)?|ease)|font-(?:mono|sans)|bg-[\w-]+|text-[\w-]+|border[\w-]*|accent[\w-]*|selection-[\w-]+|titlebar-[\w-]+|sheet-tab-[\w-]+|chart-[\w-]+|tooltip-[\w-]+|cmdk-[\w-]+|shadow-[\w-]+|formula-[\w-]+|grid-header-[\w-]+|grid-line|panel-(?:bg|border|shadow)|dialog-(?:bg|border|shadow|backdrop)|link|error[\w-]*|warning[\w-]*|success[\w-]*))\s*\2/gi;
  const tokenOpStyleBracket =
    /\[\s*(?:["'`])style(?:["'`])\s*]\s*(?:\?\.|\.)\s*(?<op>setProperty|removeProperty)\s*\(\s*(["'`])\s*(?<prop>--(?:space-\d+|radius[-\w]*|motion-(?:duration(?:-fast)?|ease)|font-(?:mono|sans)|bg-[\w-]+|text-[\w-]+|border[\w-]*|accent[\w-]*|selection-[\w-]+|titlebar-[\w-]+|sheet-tab-[\w-]+|chart-[\w-]+|tooltip-[\w-]+|cmdk-[\w-]+|shadow-[\w-]+|formula-[\w-]+|grid-header-[\w-]+|grid-line|panel-(?:bg|border|shadow)|dialog-(?:bg|border|shadow|backdrop)|link|error[\w-]*|warning[\w-]*|success[\w-]*))\s*\2/gi;
  const tokenOpStyleBracketBracket =
    /\[\s*(?:["'`])style(?:["'`])\s*]\s*(?:\?\.)?\s*\[\s*(?:["'`])(?<op>setProperty|removeProperty)(?:["'`])\s*]\s*\(\s*(["'`])\s*(?<prop>--(?:space-\d+|radius[-\w]*|motion-(?:duration(?:-fast)?|ease)|font-(?:mono|sans)|bg-[\w-]+|text-[\w-]+|border[\w-]*|accent[\w-]*|selection-[\w-]+|titlebar-[\w-]+|sheet-tab-[\w-]+|chart-[\w-]+|tooltip-[\w-]+|cmdk-[\w-]+|shadow-[\w-]+|formula-[\w-]+|grid-header-[\w-]+|grid-line|panel-(?:bg|border|shadow)|dialog-(?:bg|border|shadow|backdrop)|link|error[\w-]*|warning[\w-]*|success[\w-]*))\s*\2/gi;
  // Also guard against scripts overriding core tokens via style strings (cssText / setAttribute("style")).
  // These APIs are less common, but they can still override CSS variables at runtime.
  const tokenAssignmentInStyleString =
    /(?<prop>--(?:space-\d+|radius[-\w]*|motion-(?:duration(?:-fast)?|ease)|font-(?:mono|sans)|bg-[\w-]+|text-[\w-]+|border[\w-]*|accent[\w-]*|selection-[\w-]+|titlebar-[\w-]+|sheet-tab-[\w-]+|chart-[\w-]+|tooltip-[\w-]+|cmdk-[\w-]+|shadow-[\w-]+|formula-[\w-]+|grid-header-[\w-]+|grid-line|panel-(?:bg|border|shadow)|dialog-(?:bg|border|shadow|backdrop)|link|error[\w-]*|warning[\w-]*|success[\w-]*))\s*:/gi;
  const cssTextAssignment = /\.\s*style\b\s*\.\s*cssText\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi;
  const cssTextBracketAssignment =
    /\.\s*style\b\s*\[\s*(?:["'`])cssText(?:["'`])\s*]\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi;
  const cssTextStyleBracketAssignment =
    /\[\s*(?:["'`])style(?:["'`])\s*]\s*\.\s*cssText\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi;
  const cssTextStyleBracketBracketAssignment =
    /\[\s*(?:["'`])style(?:["'`])\s*]\s*\[\s*(?:["'`])cssText(?:["'`])\s*]\s*(?:=|\+=)\s*(["'`])\s*(?<value>[^"'`]*?)\1/gi;
  const setAttributeStyle = /\bsetAttribute\s*\(\s*(["'])style\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi;
  const setAttributeStyleBracket =
    /\[\s*(?:["'`])setAttribute(?:["'`])\s*]\s*\(\s*(["'])style\1\s*,\s*(["'`])\s*(?<value>[^"'`]*?)\2/gi;

  for (const file of files) {
    const source = fs.readFileSync(file, "utf8");
    const stripped = stripComments(source);
    const rel = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

    let match;
    for (const re of [tokenOp, tokenOpBracket, tokenOpStyleBracket, tokenOpStyleBracketBracket]) {
      re.lastIndex = 0;
      while ((match = re.exec(stripped))) {
        const prop = match.groups?.prop ?? "";
        const op = match.groups?.op ?? "setProperty";
        const matchStart = match.index ?? 0;
        const matchText = match[0] ?? "";
        const propOffset = prop ? matchText.indexOf(prop) : -1;
        const propIndex = propOffset >= 0 ? matchStart + propOffset : matchStart;
        const line = getLineNumber(stripped, propIndex);
        violations.push(`${rel}:L${line}: ${op}(${prop}, ...)`);
      }
      re.lastIndex = 0;
    }

    for (const { re, kind } of [
      { re: cssTextAssignment, kind: "style.cssText" },
      { re: cssTextBracketAssignment, kind: "style[cssText]" },
      { re: cssTextStyleBracketAssignment, kind: "style['style'].cssText" },
      { re: cssTextStyleBracketBracketAssignment, kind: "style['style'][cssText]" },
      { re: setAttributeStyle, kind: "setAttribute(style)" },
      { re: setAttributeStyleBracket, kind: "setAttribute[style]" },
    ]) {
      re.lastIndex = 0;
      while ((match = re.exec(stripped))) {
        const value = match.groups?.value ?? "";
        if (!value) continue;

        tokenAssignmentInStyleString.lastIndex = 0;
        let tokenMatch;
        while ((tokenMatch = tokenAssignmentInStyleString.exec(value))) {
          const prop = tokenMatch.groups?.prop ?? "";
          const matchStart = match.index ?? 0;
          const matchText = match[0] ?? "";
          const valueOffset = matchText.indexOf(value);
          const absIndex = matchStart + (valueOffset >= 0 ? valueOffset : 0) + (tokenMatch.index ?? 0);
          const line = getLineNumber(stripped, absIndex);
          violations.push(`${rel}:L${line}: ${kind} defines ${prop}`);
        }
        tokenAssignmentInStyleString.lastIndex = 0;
      }
      re.lastIndex = 0;
    }
  }

  assert.deepEqual(
    violations,
    [],
    `Found core token overrides in desktop UI scripts. Core design tokens must only be defined in src/styles/tokens.css:\n${violations
      .map((v) => `- ${v}`)
      .join("\n")}`,
  );
});
