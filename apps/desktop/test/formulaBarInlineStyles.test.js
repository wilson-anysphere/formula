import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments, stripCssComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("FormulaBarView avoids inline display toggles (use CSS state classes)", () => {
  const viewPath = path.join(__dirname, "..", "src", "formula-bar", "FormulaBarView.ts");
  const source = stripComments(fs.readFileSync(viewPath, "utf8"));

  assert.equal(
    source.includes(".style.display"),
    false,
    "FormulaBarView should not use element.style.display for view state; use a CSS class/attribute toggle instead",
  );

  for (const className of ["formula-bar--editing", "formula-bar--has-error", "formula-bar--error-panel-open"]) {
    assert.match(
      source,
      new RegExp(`classList\\.toggle\\(\\s*["']${className}["']`),
      `Expected FormulaBarView to toggle ${className} via classList.toggle(...)`,
    );
  }

  const cssPath = path.join(__dirname, "..", "src", "styles", "ui.css");
  const css = stripCssComments(fs.readFileSync(cssPath, "utf8"));

  for (const selector of [
    ".formula-bar:not(.formula-bar--editing) .formula-bar-input",
    ".formula-bar--has-error .formula-bar-error-button",
    ".formula-bar--has-error.formula-bar--error-panel-open .formula-bar-error-panel",
  ]) {
    assert.ok(css.includes(selector), `Expected ui.css to define ${selector}`);
  }
});
