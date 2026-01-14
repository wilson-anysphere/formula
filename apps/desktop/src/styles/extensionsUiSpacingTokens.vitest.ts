import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

import { stripCssNonSemanticText } from "../../test/testUtils/stripCssNonSemanticText.js";

function getLineNumber(text: string, index: number) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

describe("extensions-ui.css", () => {
  it("does not hardcode spacing in padding/margin/gap declarations (use --space-* tokens)", () => {
    const stylesRoot = path.dirname(fileURLToPath(import.meta.url));
    const target = path.join(stylesRoot, "extensions-ui.css");
    const css = fs.readFileSync(target, "utf8");
    const stripped = stripCssNonSemanticText(css);

    // We intentionally keep this lightweight (regex-based) to avoid pulling in a full CSS parser.
    // The goal is specifically to prevent regressions where px values creep back into spacing
    // declarations in the extensions UI overlays/dialogs (including px values hidden behind local vars).
    const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
    const spacingProp = /^(?:gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)$/i;
    const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;
    const pxUnit = /([+-]?(?:\d+(?:\.\d+)?|\.\d+))px(?![A-Za-z0-9_])/gi;

    /** @type {string[]} */
    const violations: string[] = [];
    /** @type {Set<string>} */
    const spacingVarRefs = new Set<string>();

    let decl: RegExpExecArray | null;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl.groups?.prop ?? "";
      if (!spacingProp.test(prop)) continue;

      const value = decl.groups?.value ?? "";
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      // Capture any CSS custom properties referenced by spacing declarations so we can also
      // prevent hardcoded px values from being hidden behind a local variable.
      let varMatch: RegExpExecArray | null;
      while ((varMatch = cssVarRef.exec(value))) {
        spacingVarRefs.add(varMatch[1]);
      }
      cssVarRef.lastIndex = 0;

      let unitMatch: RegExpExecArray | null;
      while ((unitMatch = pxUnit.exec(value))) {
        const numeric = unitMatch[1] ?? "";
        const n = Number(numeric);
        if (!Number.isFinite(n)) continue;
        if (n === 0) continue;

        const absIndex = valueStart + (unitMatch.index ?? 0);
        const line = getLineNumber(stripped, absIndex);
        violations.push(`${path.relative(stylesRoot, target)}:${line} ${prop}: ${value.trim()}`);
      }
      pxUnit.lastIndex = 0;
    }
    cssDeclaration.lastIndex = 0;

    // Second pass: if this file defines any custom properties that are used by spacing declarations,
    // ensure those variables also stay token-based (no hardcoded px). Include transitive references so
    // `padding: var(--a)` cannot hide `--a: var(--b)` + `--b: 8px`.
    const customPropDecls = new Map<string, Array<{ value: string; valueStart: number }>>();
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl.groups?.prop ?? "";
      if (!prop.startsWith("--")) continue;

      const value = decl.groups?.value ?? "";
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      let entries = customPropDecls.get(prop);
      if (!entries) {
        entries = [];
        customPropDecls.set(prop, entries);
      }
      entries.push({ value, valueStart });
    }
    cssDeclaration.lastIndex = 0;

    const expandedVarRefs = new Set(spacingVarRefs);
    const queue = [...spacingVarRefs];
    while (queue.length > 0) {
      const varName = queue.pop();
      if (!varName) continue;
      if (varName.startsWith("--space-")) continue;
      const declsForVar = customPropDecls.get(varName);
      if (!declsForVar) continue;

      for (const { value } of declsForVar) {
        let varMatch: RegExpExecArray | null;
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

      for (const { value, valueStart } of declsForVar) {
        let unitMatch: RegExpExecArray | null;
        while ((unitMatch = pxUnit.exec(value))) {
          const numeric = unitMatch[1] ?? "";
          const n = Number(numeric);
          if (!Number.isFinite(n)) continue;
          if (n === 0) continue;

          const absIndex = valueStart + (unitMatch.index ?? 0);
          const line = getLineNumber(stripped, absIndex);
          const raw = unitMatch[0] ?? `${numeric}px`;
          violations.push(`${path.relative(stylesRoot, target)}:${line} ${prop}: ${value.trim()} (found ${raw})`);
        }
        pxUnit.lastIndex = 0;
      }
    }

    expect(violations).toEqual([]);
  });
});
