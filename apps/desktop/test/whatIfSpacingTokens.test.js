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

test("what-if styles keep spacing on the shared --space-* scale", () => {
  const cssPath = path.join(desktopRoot, "src", "styles", "what-if.css");
  const css = fs.readFileSync(cssPath, "utf8");
  const stripped = stripCssNonSemanticText(css);

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
  const spacingProp = /^(?:gap|row-gap|column-gap|padding(?:-[a-z]+)?|margin(?:-[a-z]+)?)$/i;
  const unitRegex = /([+-]?(?:\d+(?:\.\d+)?|\.\d+))px(?![A-Za-z0-9_])/gi;

  /** @type {Set<string>} */
  const violations = new Set();

  let decl;
  while ((decl = cssDeclaration.exec(stripped))) {
    const prop = decl?.groups?.prop ?? "";
    if (!spacingProp.test(prop)) continue;

    const value = decl?.groups?.value ?? "";
    const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

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

  assert.deepEqual(
    [...violations],
    [],
    `Found pixel-based spacing declarations in what-if.css (use --space-* tokens for padding/margin/gap):\n${[...violations]
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );
});
