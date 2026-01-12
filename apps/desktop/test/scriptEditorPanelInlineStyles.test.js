import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

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
  const source = fs.readFileSync(panelPath, "utf8");

  const mountSection = extractSection(source, "export function mountScriptEditorPanel", "function defaultScript()");

  assert.equal(
    /\.style\./.test(mountSection),
    false,
    "mountScriptEditorPanel should not set inline styles; use token-based CSS classes instead",
  );
  assert.equal(
    /setAttribute\(\s*["']style["']/.test(mountSection),
    false,
    "mountScriptEditorPanel should not set inline styles via setAttribute('style', ...); use token-based CSS classes instead",
  );

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

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSrc = fs.readFileSync(mainPath, "utf8");
  assert.equal(
    /import\s+["'][^"']*styles\/script-editor\.css["']/.test(mainSrc),
    true,
    "apps/desktop/src/main.ts should import src/styles/script-editor.css so the Script Editor panel is styled in production builds",
  );
});
