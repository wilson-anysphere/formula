import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("Script Editor panel scaffold is class-driven (no inline style assignments)", () => {
  const srcPath = path.join(__dirname, "..", "src", "panels", "script-editor", "ScriptEditorPanel.js");
  const src = fs.readFileSync(srcPath, "utf8");

  const forbidden = [
    // Toolbar styles.
    "toolbar.style.display",
    "toolbar.style.gap",
    "toolbar.style.padding",
    "toolbar.style.borderBottom",
    // Editor host styles.
    "editorHost.style.flex",
    "editorHost.style.minHeight",
    // Console styles.
    "consoleHost.style.height",
    "consoleHost.style.margin",
    "consoleHost.style.padding",
    "consoleHost.style.overflow",
    "consoleHost.style.borderTop",
    // Root styles.
    "root.style.display",
    "root.style.flexDirection",
    "root.style.height",
    // Fallback editor styles.
    "fallbackEditor.style.width",
    "fallbackEditor.style.height",
    "fallbackEditor.style.resize",
    "fallbackEditor.style.border",
    "fallbackEditor.style.outline",
    "fallbackEditor.style.padding",
    "fallbackEditor.style.boxSizing",
    "fallbackEditor.style.fontFamily",
    "fallbackEditor.style.fontSize",
  ];

  for (const snippet of forbidden) {
    assert.equal(
      src.includes(snippet),
      false,
      `ScriptEditorPanel.js should not set inline styles (found: ${snippet}); use src/styles/script-editor.css classes instead`,
    );
  }

  assert.equal(
    /\.style\b/.test(src) || /setAttribute\(\s*["']style["']/.test(src),
    false,
    "ScriptEditorPanel.js should not set inline styles (element.style* / setAttribute('style', ...)); use src/styles/script-editor.css classes instead",
  );

  const requiredClasses = [
    "script-editor",
    "script-editor__toolbar",
    "script-editor__editor-host",
    "script-editor__console",
    "script-editor__fallback-editor",
  ];
  for (const className of requiredClasses) {
    assert.ok(src.includes(className), `Expected ScriptEditorPanel.js to reference CSS class "${className}"`);
  }

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSrc = fs.readFileSync(mainPath, "utf8");
  assert.equal(
    /import\s+["'][^"']*styles\/script-editor\.css["']/.test(mainSrc),
    true,
    "apps/desktop/src/main.ts should import src/styles/script-editor.css so the Script Editor panel is styled in production builds",
  );
});

