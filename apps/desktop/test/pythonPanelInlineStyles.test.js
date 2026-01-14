import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("PythonPanel is class-driven (no static inline style assignments)", () => {
  const srcPath = path.join(__dirname, "..", "src", "panels", "python", "PythonPanel.tsx");
  const src = fs.readFileSync(srcPath, "utf8");

  assert.equal(
    /\.style\b/.test(src) || /setAttribute\(\s*["']style["']/.test(src),
    false,
    "PythonPanel should not use inline styles (element.style* / setAttribute('style', ...)); use src/styles/python-panel.css classes instead",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSrc = stripComments(fs.readFileSync(mainPath, "utf8"));
  assert.equal(
    /^\s*import\s+["'][^"']*styles\/python-panel\.css["']\s*;?/m.test(mainSrc),
    true,
    "apps/desktop/src/main.ts should import src/styles/python-panel.css so the Python panel is styled in production builds",
  );

  const requiredClasses = [
    "python-panel",
    "python-panel__toolbar",
    "python-panel__network-label",
    "python-panel__allowlist-input",
    "python-panel__split",
    "python-panel__editor-wrap",
    "python-panel__editor",
    "python-panel__output",
  ];

  for (const className of requiredClasses) {
    assert.ok(src.includes(className), `Expected PythonPanel.tsx to reference CSS class "${className}"`);
  }
});

test("pythonPanelMount is class-driven (no static inline style assignments)", () => {
  const srcPath = path.join(__dirname, "..", "src", "panels", "python", "pythonPanelMount.js");
  const src = fs.readFileSync(srcPath, "utf8");

  assert.equal(
    /\.style\b/.test(src) || /setAttribute\(\s*["']style["']/.test(src),
    false,
    "pythonPanelMount should not use inline styles (element.style* / setAttribute('style', ...)); use src/styles/python-panel.css classes instead",
  );

  const requiredClasses = [
    "python-panel-mount",
    "python-panel-mount__toolbar",
    "python-panel-mount__runtime-select",
    "python-panel-mount__isolation-label",
    "python-panel-mount__degraded-banner",
    "python-panel-mount__editor-host",
    "python-panel-mount__editor",
    "python-panel-mount__console",
  ];

  for (const className of requiredClasses) {
    assert.ok(src.includes(className), `Expected pythonPanelMount.js to reference CSS class "${className}"`);
  }
});

test("python-panel.css defines the pythonPanelMount class selectors", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "python-panel.css");
  const css = fs.readFileSync(cssPath, "utf8");

  const requiredSelectors = [
    "python-panel-mount",
    "python-panel-mount__toolbar",
    "python-panel-mount__runtime-select",
    "python-panel-mount__isolation-label",
    "python-panel-mount__degraded-banner",
    "python-panel-mount__editor-host",
    "python-panel-mount__editor",
    "python-panel-mount__console",
  ];

  for (const className of requiredSelectors) {
    assert.ok(css.includes(`.${className}`), `Expected python-panel.css to define selector ".${className}"`);
  }
});
