import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("what-if styles keep spacing on the shared --space-* scale (no hardcoded lengths)", () => {
  const cssPath = path.join(desktopRoot, "src", "styles", "what-if.css");
  const css = fs.readFileSync(cssPath, "utf8");
  const stripped = stripCssNonSemanticText(css);

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
  const spacingProp = /^(?:gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)$/i;
  const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;
  const unitRegex =
    /([+-]?(?:\d+(?:\.\d+)?|\.\d+))(px|%|rem|em|vh|vw|vmin|vmax|cm|mm|in|pt|pc|ch|ex)(?![A-Za-z0-9_])/gi;

  /** @type {Set<string>} */
  const violations = new Set();
  /** @type {Set<string>} */
  const spacingVarRefs = new Set();

  let decl;
  while ((decl = cssDeclaration.exec(stripped))) {
    const prop = decl?.groups?.prop ?? "";
    if (!spacingProp.test(prop)) continue;

    const value = decl?.groups?.value ?? "";
    const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

    // Capture any CSS custom properties referenced by spacing declarations so we can also
    // prevent hardcoded units from being hidden behind a local variable.
    let varMatch;
    while ((varMatch = cssVarRef.exec(value))) {
      spacingVarRefs.add(varMatch[1]);
    }
    cssVarRef.lastIndex = 0;

    let unitMatch;
    while ((unitMatch = unitRegex.exec(value))) {
      const numeric = unitMatch[1];
      const n = Number(numeric);
      if (!Number.isFinite(n)) continue;
      if (n === 0) continue;

      const absIndex = valueStart + unitMatch.index;
      const line = getLineNumber(stripped, absIndex);
      const rel = path.relative(desktopRoot, cssPath).replace(/\\\\/g, "/");
      violations.add(`${rel}:L${line}: ${prop}: ${value.trim()}`);
    }

    unitRegex.lastIndex = 0;
  }

  // Second pass: if this file defines any custom properties that are used by spacing declarations,
  // ensure those variables also stay token-based (no hardcoded units). Include transitive references so
  // `padding: var(--a)` cannot hide `--a: var(--b)` + `--b: 8px`.
  /** @type {Map<string, Array<{ value: string, valueStart: number }>>} */
  const customPropDecls = new Map();

  cssDeclaration.lastIndex = 0;
  while ((decl = cssDeclaration.exec(stripped))) {
    const prop = decl?.groups?.prop ?? "";
    if (!prop.startsWith("--")) continue;

    const value = decl?.groups?.value ?? "";
    const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

    let entries = customPropDecls.get(prop);
    if (!entries) {
      entries = [];
      customPropDecls.set(prop, entries);
    }
    entries.push({ value, valueStart });
  }
  cssDeclaration.lastIndex = 0;

  /** @type {Set<string>} */
  const expandedSpacingVarRefs = new Set(spacingVarRefs);
  const queue = [...spacingVarRefs];
  while (queue.length > 0) {
    const varName = queue.pop();
    const declsForVar = customPropDecls.get(varName);
    if (!declsForVar) continue;

    for (const { value } of declsForVar) {
      let varMatch;
      while ((varMatch = cssVarRef.exec(value))) {
        const ref = varMatch[1];
        if (expandedSpacingVarRefs.has(ref)) continue;
        expandedSpacingVarRefs.add(ref);
        queue.push(ref);
      }
      cssVarRef.lastIndex = 0;
    }
  }

  for (const [prop, declsForVar] of customPropDecls) {
    if (!expandedSpacingVarRefs.has(prop)) continue;

    for (const { value, valueStart } of declsForVar) {
      let unitMatch;
      while ((unitMatch = unitRegex.exec(value))) {
        const numeric = unitMatch[1];
        const n = Number(numeric);
        if (!Number.isFinite(n)) continue;
        if (n === 0) continue;

        const absIndex = valueStart + unitMatch.index;
        const line = getLineNumber(stripped, absIndex);
        const rel = path.relative(desktopRoot, cssPath).replace(/\\\\/g, "/");
        violations.add(`${rel}:L${line}: ${prop}: ${value.trim()}`);
      }

      unitRegex.lastIndex = 0;
    }
  }

  assert.deepEqual(
    [...violations],
    [],
    `Found hardcoded spacing values in what-if.css (use --space-* tokens for padding/margin/gap):\n${[...violations]
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});
