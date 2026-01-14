import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments, stripCssComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("renderMacroRunner is class-driven (no inline style assignments)", () => {
  const srcPath = path.join(__dirname, "..", "src", "macros", "dom_ui.ts");
  const src = stripComments(fs.readFileSync(srcPath, "utf8"));

  const usesInlineStyle =
    // Direct DOM style manipulation (element.style.* / element.style = ... / etc).
    /\.style\b/.test(src) ||
    // Bracket access to the style property (element["style"] / element['style']).
    /\[\s*["']style["']\s*\]/.test(src) ||
    // Attribute-based inline styles.
    /setAttribute\(\s*["']style["']/.test(src) ||
    /setAttributeNS\(\s*[^,]+,\s*["']style["']/.test(src) ||
    // Inline style attributes in HTML strings (e.g. element.innerHTML = "<div style=\"...\">").
    /<[^>]*\bstyle\s*=\s*["']/i.test(src);

  assert.equal(
    usesInlineStyle,
    false,
    "renderMacroRunner should not use inline styles (element.style* / setAttribute('style', ...)); use src/styles/macros-runner.css classes instead",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSrc = stripComments(fs.readFileSync(mainPath, "utf8"));
  assert.equal(
    /^\s*import\s+["'][^"']*styles\/macros-runner\.css["']\s*;?/m.test(mainSrc),
    true,
    "apps/desktop/src/main.ts should import src/styles/macros-runner.css so the macro runner UI is styled in production builds",
  );

  const cssPath = path.join(__dirname, "..", "src", "styles", "macros-runner.css");
  assert.equal(
    fs.existsSync(cssPath),
    true,
    "Expected apps/desktop/src/styles/macros-runner.css to exist (macro runner styling should live in a dedicated stylesheet)",
  );
  const css = stripCssComments(fs.readFileSync(cssPath, "utf8"));
  assert.match(css, /\.macros-runner\b/, "Expected macros-runner.css to define a .macros-runner selector");
  assert.match(
    css,
    /\.macros-runner__controls\s*\{[^}]*\bdisplay\s*:\s*flex\b[^}]*\}/,
    "Expected .macros-runner__controls to lay out the select + buttons as a flex row",
  );

  const requiredClasses = [
    "macros-runner",
    "macros-runner__header",
    "macros-runner__security-banner",
    "macros-runner__controls",
    "macros-runner__select",
    "macros-runner__button",
    "macros-runner__output",
  ];

  for (const className of requiredClasses) {
    assert.ok(src.includes(className), `Expected dom_ui.ts to reference CSS class "${className}"`);
    assert.ok(
      css.includes(`.${className}`),
      `Expected macros-runner.css to define CSS for ".${className}"`,
    );
  }
});
