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

test("desktop UI should not hardcode border-radius values (use radius tokens)", () => {
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
  /** @type {Set<string>} */
  const globalBorderRadiusVarRefs = new Set();

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
  const borderRadiusProp = /^border(?:-(?:top|bottom|start|end)-(?:left|right|start|end))?-radius$/i;
  const cssVarRef = /\bvar\(\s*(--[-\w]+)\b/g;
  const unitRegex = /([+-]?(?:\d+(?:\.\d+)?|\.\d+))(px|%|rem|em|vh|vw|vmin|vmax|cm|mm|in|pt|pc|ch|ex)(?![A-Za-z0-9_])/gi;

  for (const file of files) {
    const css = fs.readFileSync(file, "utf8");
    const stripped = stripCssNonSemanticText(css);

    /** @type {Set<string>} */
    const borderRadiusVarRefs = new Set();

    let decl;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl?.groups?.prop ?? "";
      if (!borderRadiusProp.test(prop)) continue;

      const value = decl?.groups?.value ?? "";
      // `decl[0]` ends with the captured group, so this points at the first character of the value.
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      // Capture any CSS custom properties referenced by border-radius declarations so we can also
      // prevent hardcoded units from being hidden behind a local variable.
      let varMatch;
      while ((varMatch = cssVarRef.exec(value))) {
        borderRadiusVarRefs.add(varMatch[1]);
        globalBorderRadiusVarRefs.add(varMatch[1]);
      }
      cssVarRef.lastIndex = 0;

      let unitMatch;
      while ((unitMatch = unitRegex.exec(value))) {
        const numeric = unitMatch[1];
        const unit = unitMatch[2] ?? "";
        const n = Number(numeric);
        if (!Number.isFinite(n)) continue;
        if (n === 0) continue;

        const absIndex = valueStart + unitMatch.index;
        const line = getLineNumber(stripped, absIndex);
        violations.push(
          `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: border-radius: ${numeric}${unit}`,
        );
      }

      unitRegex.lastIndex = 0;
    }

    // Second pass: if this file defines any custom properties that are used by border-radius declarations,
    // ensure those variables also stay token-based (no hardcoded units).
    cssDeclaration.lastIndex = 0;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl?.groups?.prop ?? "";
      if (!prop.startsWith("--")) continue;
      if (!borderRadiusVarRefs.has(prop)) continue;

      const value = decl?.groups?.value ?? "";
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      let unitMatch;
      while ((unitMatch = unitRegex.exec(value))) {
        const numeric = unitMatch[1];
        const unit = unitMatch[2] ?? "";
        const n = Number(numeric);
        if (!Number.isFinite(n)) continue;
        if (n === 0) continue;

        const absIndex = valueStart + unitMatch.index;
        const line = getLineNumber(stripped, absIndex);
        violations.push(
          `${path.relative(desktopRoot, file).replace(/\\\\/g, "/")}:L${line}: ${prop}: ${numeric}${unit}`,
        );
      }

      unitRegex.lastIndex = 0;
    }

    cssDeclaration.lastIndex = 0;
  }

  // Finally, ensure that any border-radius variables sourced from design tokens do not hide
  // hardcoded unit lengths (except the canonical `--radius*` tokens themselves).
  const tokensCssPath = path.join(srcRoot, "styles", "tokens.css");
  if (fs.existsSync(tokensCssPath)) {
    const css = fs.readFileSync(tokensCssPath, "utf8");
    const stripped = stripCssNonSemanticText(css);

    let decl;
    while ((decl = cssDeclaration.exec(stripped))) {
      const prop = decl?.groups?.prop ?? "";
      if (!prop.startsWith("--")) continue;
      if (!globalBorderRadiusVarRefs.has(prop)) continue;
      if (prop.startsWith("--radius")) continue;

      const value = decl?.groups?.value ?? "";
      const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

      let unitMatch;
      while ((unitMatch = unitRegex.exec(value))) {
        const numeric = unitMatch[1];
        const unit = unitMatch[2] ?? "";
        const n = Number(numeric);
        if (!Number.isFinite(n)) continue;
        if (n === 0) continue;

        const absIndex = valueStart + unitMatch.index;
        const line = getLineNumber(stripped, absIndex);
        const rawUnit = unitMatch[0] ?? `${numeric}${unit}`;
        violations.push(
          `${path.relative(desktopRoot, tokensCssPath).replace(/\\\\/g, "/")}:L${line}: ${prop}: ${value.trim()} (found ${rawUnit})`,
        );
      }

      unitRegex.lastIndex = 0;
    }

    cssDeclaration.lastIndex = 0;
  }

  assert.deepEqual(
    violations,
    [],
    `Found hardcoded border-radius values in desktop UI styles. Use radius tokens (var(--radius*)), except for 0:\n${violations
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});
