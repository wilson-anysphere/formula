import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

import { stripComments, stripCssComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function extractSection(source, startMarker, endMarker) {
  const startIdx = source.indexOf(startMarker);
  assert.ok(startIdx !== -1, `Expected to find start marker: ${startMarker}`);

  const endIdx = source.indexOf(endMarker, startIdx);
  assert.ok(endIdx !== -1, `Expected to find end marker: ${endMarker}`);

  return source.slice(startIdx, endIdx);
}

test("ScriptEditorPanel avoids static inline styles and uses token-based classes", () => {
  const panelPath = path.join(__dirname, "..", "src", "panels", "script-editor", "ScriptEditorPanel.js");
  const source = stripComments(fs.readFileSync(panelPath, "utf8"));

  assert.equal(
    /\.style\b/.test(source) || /setAttribute\(\s*["']style["']/.test(source),
    false,
    "ScriptEditorPanel.js should not use inline styles (element.style* / setAttribute('style', ...)); use src/styles/script-editor.css classes instead",
  );

  const mountSection = extractSection(source, "export function mountScriptEditorPanel", "function defaultScript()");

  assert.match(
    mountSection,
    /root\.className\s*=\s*"script-editor"/,
    "mountScriptEditorPanel should apply the script-editor root class",
  );
  assert.match(
    mountSection,
    /toolbar\.className\s*=\s*"script-editor__toolbar"/,
    "mountScriptEditorPanel should apply the script-editor__toolbar class",
  );
  assert.match(
    mountSection,
    /runButton\.className\s*=\s*"script-editor__run-button"/,
    "mountScriptEditorPanel should apply the script-editor__run-button class",
  );
  assert.match(
    mountSection,
    /editorHost\.className\s*=\s*"script-editor__editor-host"/,
    "mountScriptEditorPanel should apply the script-editor__editor-host class",
  );
  assert.match(
    mountSection,
    /consoleHost\.className\s*=\s*"script-editor__console"/,
    "mountScriptEditorPanel should apply the script-editor__console class",
  );
  assert.match(
    mountSection,
    /fallbackEditor\.className\s*=\s*"script-editor__fallback-editor"/,
    "mountScriptEditorPanel should apply the script-editor__fallback-editor class",
  );
  assert.match(
    mountSection,
    /runButton\.dataset\.testid\s*=\s*"script-editor-run"/,
    "mountScriptEditorPanel should preserve data-testid=\"script-editor-run\"",
  );
  assert.match(
    mountSection,
    /fallbackEditor\.dataset\.testid\s*=\s*"script-editor-code"/,
    "mountScriptEditorPanel should preserve data-testid=\"script-editor-code\"",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSrc = stripComments(fs.readFileSync(mainPath, "utf8"));
  assert.equal(
    /^\s*import\s+["'][^"']*styles\/script-editor\.css["']\s*;?/m.test(mainSrc),
    true,
    "apps/desktop/src/main.ts should import src/styles/script-editor.css so the Script Editor panel is styled in production builds",
  );

  const cssPath = path.join(__dirname, "..", "src", "styles", "script-editor.css");
  const css = stripCssComments(fs.readFileSync(cssPath, "utf8"));
  assert.match(css, /\.script-editor\s*\{/);
  assert.match(css, /\.script-editor__toolbar\s*\{/);
  assert.match(css, /\.script-editor__run-button\s*\{/);
  assert.match(css, /\.script-editor__editor-host\s*\{/);
  assert.match(css, /\.script-editor__editor-host\s*\{[^}]*\bflex:\s*1\b[^}]*\}/s);
  assert.match(css, /\.script-editor__editor-host\s*\{[^}]*\bmin-height:\s*0\b[^}]*\}/s);

  // Sanity-check that the fallback editor is styled as a code editor (monospace) and non-resizable.
  assert.match(css, /\.script-editor__fallback-editor\s*\{/);
  assert.match(
    css,
    /\.script-editor__fallback-editor\s*\{[^}]*\b(?:font-family|font)\s*:[^;]*var\(--font-mono\)[^;]*;[^}]*\}/s,
  );
  assert.match(css, /\.script-editor__fallback-editor\s*\{[^}]*\bresize:\s*none\b[^}]*\}/s);

  // Preserve the console sizing/overflow behavior (important for output visibility).
  assert.match(css, /\.script-editor__console\s*\{/);
  assert.match(css, /\.script-editor__console\s*\{[^}]*\bheight:\s*140px\b[^}]*\}/s);
  assert.match(css, /\.script-editor__console\s*\{[^}]*\boverflow:\s*auto\b[^}]*\}/s);
});
