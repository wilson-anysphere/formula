import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripCssNonSemanticText } from "./testUtils/stripCssNonSemanticText.js";
import { stripComments, stripCssComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function getLineNumber(text, index) {
  return text.slice(0, Math.max(0, index)).split("\n").length;
}

test("Solver panel React components avoid inline styles (use solver.css classes)", () => {
  const panelDir = path.join(__dirname, "..", "src", "panels", "solver");
  const cssPath = path.join(__dirname, "..", "src", "styles", "solver.css");
  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const desktopRoot = path.join(__dirname, "..");

  const sources = {
    panel: stripComments(fs.readFileSync(path.join(panelDir, "SolverPanel.tsx"), "utf8")),
    dialog: stripComments(fs.readFileSync(path.join(panelDir, "SolverDialog.tsx"), "utf8")),
    progress: stripComments(fs.readFileSync(path.join(panelDir, "SolverProgress.tsx"), "utf8")),
    summary: stripComments(fs.readFileSync(path.join(panelDir, "SolverResultSummary.tsx"), "utf8")),
  };

  for (const [name, content] of Object.entries(sources)) {
    assert.equal(
      /\bstyle\s*=/.test(content),
      false,
      `${name} should not use inline styles; use src/styles/solver.css classes instead`,
    );
  }

  const requiredClasses = [
    "solver-panel",
    "solver-panel__header",
    "solver-panel__configure-button",
    "solver-panel__error",
    "solver-dialog",
    "solver-dialog__title",
    "solver-dialog__section",
    "solver-dialog__field",
    "solver-dialog__label",
    "solver-dialog__input",
    "solver-dialog__select",
    "solver-dialog__row",
    "solver-dialog__variable-row",
    "solver-dialog__constraint-row",
    "solver-dialog__footer",
    "solver-progress",
    "solver-results",
    "solver-results__actions",
    "solver__button",
    "solver__button--primary",
  ];

  for (const className of requiredClasses) {
    assert.ok(
      Object.values(sources).some((src) => src.includes(className)),
      `Expected solver components to reference the ${className} CSS class`,
    );
  }

  assert.equal(fs.existsSync(cssPath), true, "Expected apps/desktop/src/styles/solver.css to exist");
  const css = stripCssComments(fs.readFileSync(cssPath, "utf8"));
  const strippedCss = stripCssNonSemanticText(css);
  for (const className of requiredClasses) {
    assert.ok(css.includes(`.${className}`), `Expected solver.css to define .${className}`);
  }

  const relCssPath = path.relative(desktopRoot, cssPath).replace(/\\\\/g, "/");
  const monospaceStack = /\b(?:ui-monospace|sf\s*mono|menlo|consolas|monospace)\b/gi;
  /** @type {Set<string>} */
  const monospaceViolations = new Set();
  let monoMatch;
  while ((monoMatch = monospaceStack.exec(strippedCss))) {
    const absIndex = monoMatch.index ?? 0;
    const line = getLineNumber(strippedCss, absIndex);
    monospaceViolations.add(`${relCssPath}:L${line}: ${monoMatch[0]}`);
  }

  assert.deepEqual(
    [...monospaceViolations],
    [],
    `solver.css should not hardcode a monospace font stack; use var(--font-mono) instead:\n${[
      ...monospaceViolations,
    ]
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );

  const cssDeclaration = /(?:^|[;{])\s*(?<prop>[-\w]+)\s*:\s*(?<value>[^;{}]*)/gi;
  const spacingProp = /^(?:gap|row-gap|column-gap|padding(?:-[a-z]+)*|margin(?:-[a-z]+)*)$/i;
  const pxUnit = /([+-]?(?:\d+(?:\.\d+)?|\.\d+))px(?![A-Za-z0-9_])/gi;

  /** @type {Set<string>} */
  const pxSpacingViolations = new Set();
  let decl;
  while ((decl = cssDeclaration.exec(strippedCss))) {
    const prop = decl?.groups?.prop ?? "";
    if (!spacingProp.test(prop)) continue;

    const value = decl?.groups?.value ?? "";
    // `decl[0]` ends with the captured group, so this points at the first character of the value.
    const valueStart = (decl.index ?? 0) + decl[0].length - value.length;

    let unitMatch;
    while ((unitMatch = pxUnit.exec(value))) {
      const numeric = unitMatch[1] ?? "";
      const n = Number(numeric);
      if (!Number.isFinite(n)) continue;
      if (n === 0) continue;

      const absIndex = valueStart + (unitMatch.index ?? 0);
      const line = getLineNumber(strippedCss, absIndex);
      pxSpacingViolations.add(`${relCssPath}:L${line}: ${prop}: ${value.trim()}`);
    }

    pxUnit.lastIndex = 0;
  }

  assert.deepEqual(
    [...pxSpacingViolations],
    [],
    `solver.css should not use raw px values for layout spacing (padding*/margin*/gap); use --space-* tokens instead:\n${[
      ...pxSpacingViolations,
    ]
      .map((violation) => `- ${violation}`)
      .join("\n")}`,
  );

  const mainSrc = stripComments(fs.readFileSync(mainPath, "utf8"));
  assert.match(
    mainSrc,
    /^\s*import\s+["'][^"']*styles\/solver\.css["']\s*;?/m,
    "apps/desktop/src/main.ts should import src/styles/solver.css so the Solver panel UI is styled in production builds",
  );
});
