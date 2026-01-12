import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("renderMacroRunner is class-driven (no inline style assignments)", () => {
  const srcPath = path.join(__dirname, "..", "src", "macros", "dom_ui.ts");
  const src = fs.readFileSync(srcPath, "utf8");

  assert.equal(
    /\.style\b/.test(src) || /setAttribute\(\s*["']style["']/.test(src),
    false,
    "renderMacroRunner should not use inline styles (element.style* / setAttribute('style', ...)); use src/styles/macros-runner.css classes instead",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSrc = fs.readFileSync(mainPath, "utf8");
  assert.equal(
    /import\s+["'][^"']*styles\/macros-runner\.css["']/.test(mainSrc),
    true,
    "apps/desktop/src/main.ts should import src/styles/macros-runner.css so the macro runner UI is styled in production builds",
  );

  const cssPath = path.join(__dirname, "..", "src", "styles", "macros-runner.css");
  assert.equal(
    fs.existsSync(cssPath),
    true,
    "Expected apps/desktop/src/styles/macros-runner.css to exist (macro runner styling should live in a dedicated stylesheet)",
  );
  const css = fs.readFileSync(cssPath, "utf8");
  assert.match(css, /\.macros-runner\b/, "Expected macros-runner.css to define a .macros-runner selector");

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
