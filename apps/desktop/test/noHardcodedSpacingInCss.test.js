import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.join(__dirname, "..");
const srcRoot = path.join(desktopRoot, "src");

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

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("desktop UI CSS keeps layout spacing on the --space-* scale (no raw px in padding/margin/gap)", () => {
  const files = walkCssFiles(srcRoot).filter((file) => {
    const rel = path.relative(srcRoot, file).replace(/\\\\/g, "/");
    // Demo/sandbox assets are not part of the shipped UI bundle.
    if (rel.startsWith("grid/presence-renderer/")) return false;
    if (rel.includes("/demo/")) return false;
    if (rel.includes("/__tests__/")) return false;
    return true;
  });

  /** @type {string[]} */
  const violations = [];

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
  const spacingProp = /^(?:gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)$/i;
  const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;
  const pxUnit = /([+-]?(?:\d+(?:\.\d+)?|\.\d+))px(?![A-Za-z0-9_])/gi;

  for (const file of files) {
    const css = fs.readFileSync(file, "utf8");
    const stripped = stripCssNonSemanticText(css);

    /** @type {Set<string>} */
    const spacingVarRefs = new Set();

    let decl;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl?.groups?.prop ?? "";
      if (!spacingProp.test(prop)) continue;

      const value = decl?.groups?.value ?? "";
      // `decl[0]` ends with the captured group, so this points at the first character of the value.
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      // Capture any CSS custom properties referenced by spacing declarations so we can also
      // prevent hardcoded units from being hidden behind a local variable.
      let varMatch;
      while ((varMatch = cssVarRef.exec(value))) {
        spacingVarRefs.add(varMatch[1]);
      }
      cssVarRef.lastIndex = 0;

      let unitMatch;
      while ((unitMatch = pxUnit.exec(value))) {
        const numeric = unitMatch[1] ?? "";
        const n = Number(numeric);
        if (!Number.isFinite(n)) continue;
        if (n === 0) continue;

        const absIndex = valueStart + (unitMatch.index ?? 0);
        const line = getLineNumber(stripped, absIndex);
        violations.push(`${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${prop}: ${value.trim()}`);
      }

      pxUnit.lastIndex = 0;
    }

    // Second pass: if this file defines any custom properties that are used by spacing declarations,
    // ensure those variables also stay token-based (no hardcoded px).
    cssDeclaration.lastIndex = 0;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl?.groups?.prop ?? "";
      if (!prop.startsWith("--")) continue;
      if (!spacingVarRefs.has(prop)) continue;

      const value = decl?.groups?.value ?? "";
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      let unitMatch;
      while ((unitMatch = pxUnit.exec(value))) {
        const numeric = unitMatch[1] ?? "";
        const n = Number(numeric);
       if (!Number.isFinite(n)) continue;
       if (n === 0) continue;

       const absIndex = valueStart + (unitMatch.index ?? 0);
       const line = getLineNumber(stripped, absIndex);
        // Include the specific offending literal so multi-value declarations like
        // `padding: 8px 10px` report both 8px and 10px distinctly.
        const rawUnit = unitMatch[0] ?? `${numeric}px`;
        violations.push(`${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${prop}: ${value.trim()} (found ${rawUnit})`);
      }

      pxUnit.lastIndex = 0;
    }

    cssDeclaration.lastIndex = 0;
  }

  assert.deepEqual(
    violations,
    [],
    `Found raw px values in desktop CSS spacing properties (use --space-* tokens for padding/margin/gap; allow 0):\n${violations
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});
