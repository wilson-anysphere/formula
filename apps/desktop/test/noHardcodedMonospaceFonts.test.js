import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const srcRoot = path.join(desktopRoot, "src");

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("desktop styles should not hardcode monospace font stacks (use --font-mono token)", () => {
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

  const cssFiles = walkCssFiles(srcRoot)
    .filter((file) => {
      const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
      if (rel === "styles/tokens.css") return false;
      // Demo/sandbox assets are not part of the shipped UI bundle.
      if (rel.startsWith("grid/presence-renderer/")) return false;
      if (rel.includes("/demo/")) return false;
      if (rel.includes("/__tests__/")) return false;
      return true;
    })
    .sort((a, b) => a.localeCompare(b));

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
  const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;

  // Keep in sync with `--font-mono` in `src/styles/tokens.css`; we only allow hardcoded
  // monospace stacks in that single source of truth.
  //
  // Note: We intentionally scan only `font` / `font-family` declarations so we can preserve
  // quoted font family names like `"SF Mono"` without triggering false positives in unrelated
  // strings (e.g. `content: "monospace"` or urls).
  const forbiddenFontStackToken = /\b(?:ui-monospace|SFMono(?:-Regular)?|SF\s*Mono|Menlo|Consolas|monospace)\b/gi;
  const allowedFontVars = new Set(["--font-mono", "--font-sans"]);

  /** @type {string[]} */
  const violations = new Set();
  /** @type {Set<string>} */
  const fontVarRefs = new Set();
  /** @type {Map<string, string>} */
  const strippedByFile = new Map();
  /** @type {Map<string, Array<{ file: string, value: string, valueStart: number }>>} */
  const customPropDecls = new Map();

  for (const file of cssFiles) {
    const raw = fs.readFileSync(file, "utf8");
    const stripped = stripCssComments(raw);
    strippedByFile.set(file, stripped);
    const relPath = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

    cssDeclaration.lastIndex = 0;
    let decl;
    while ((decl = cssDeclaration.exec(stripped))) {
      const propRaw = decl?.groups?.prop ?? "";
      const prop = propRaw.toLowerCase();

      const value = decl?.groups?.value ?? "";
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      if (propRaw.startsWith("--")) {
        let entries = customPropDecls.get(propRaw);
        if (!entries) {
          entries = [];
          customPropDecls.set(propRaw, entries);
        }
        entries.push({ file, value, valueStart });
      }

      if (prop !== "font-family" && prop !== "font") continue;

      let varMatch;
      while ((varMatch = cssVarRef.exec(value))) {
        fontVarRefs.add(varMatch[1]);
      }
      cssVarRef.lastIndex = 0;

      forbiddenFontStackToken.lastIndex = 0;
      let match;
      while ((match = forbiddenFontStackToken.exec(value))) {
        const absIndex = valueStart + (match.index ?? 0);
        const line = getLineNumber(stripped, absIndex);
        violations.add(`${relPath}:L${line}: ${match[0]}`);
      }
    }
    forbiddenFontStackToken.lastIndex = 0;
  }

  // Ensure that any font variables referenced by font/font-family declarations do not hide
  // hardcoded monospace stacks, even if those variables are defined in a different stylesheet.
  // Include transitive references so `font-family: var(--a)` cannot hide `--a: var(--b)` + `--b: monospace`.
  /** @type {Set<string>} */
  const expandedFontVarRefs = new Set(fontVarRefs);
  const queue = [...fontVarRefs];
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
        if (expandedFontVarRefs.has(ref)) continue;
        expandedFontVarRefs.add(ref);
        queue.push(ref);
      }
      cssVarRef.lastIndex = 0;
    }
  }

  for (const [prop, declsForVar] of customPropDecls) {
    if (!expandedFontVarRefs.has(prop)) continue;
    if (allowedFontVars.has(prop)) continue;

    for (const { file, value, valueStart } of declsForVar) {
      const stripped = strippedByFile.get(file) ?? "";
      const relPath = path.relative(desktopRoot, file).replace(/\\\\/g, "/");

      forbiddenFontStackToken.lastIndex = 0;
      let match;
      while ((match = forbiddenFontStackToken.exec(value))) {
        const absIndex = valueStart + (match.index ?? 0);
        const line = getLineNumber(stripped, absIndex);
        violations.add(`${relPath}:L${line}: ${prop}: ${value.trim()} (found ${match[0]})`);
      }
    }
  }

  assert.deepEqual(
    [...violations],
    [],
    `Found hardcoded monospace font stacks in desktop styles. Use var(--font-mono) from src/styles/tokens.css:\n${[
      ...violations,
    ]
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});
