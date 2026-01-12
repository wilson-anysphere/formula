import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

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
  const mainSrc = fs.readFileSync(mainPath, "utf8");
  assert.equal(
    /import\s+["'][^"']*styles\/python-panel\.css["']/.test(mainSrc),
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

