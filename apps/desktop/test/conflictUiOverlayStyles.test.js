import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");

function extractSection(source, startMarker, endMarker) {
  const startIdx = source.indexOf(startMarker);
  assert.ok(startIdx !== -1, `Expected to find start marker: ${startMarker}`);

  if (!endMarker) return source.slice(startIdx);

  const endIdx = source.indexOf(endMarker, startIdx);
  assert.ok(endIdx !== -1, `Expected to find end marker: ${endMarker}`);

  return source.slice(startIdx, endIdx);
}

test("spreadsheet conflict UI overlay uses CSS classes (no static inline styles)", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  const conflictUiSection = extractSection(
    text,
    "// Conflicts UI (mounted once; new conflicts stream in via the monitor callbacks).",
    "const presence = this.collabSession.presence;",
  );

  assert.equal(
    /\.style\b/.test(conflictUiSection),
    false,
    "expected conflict overlay block to avoid inline style.*; presentation should be CSS-driven",
  );

  // Conflict UI overlay container: presentation styling should be CSS-driven.
  assert.ok(
    !text.includes("this.conflictUiContainer.style.position"),
    "expected conflict overlay position to be driven by CSS",
  );
  assert.ok(!text.includes("this.conflictUiContainer.style.inset"), "expected conflict overlay inset to be driven by CSS");
  assert.ok(!text.includes("this.conflictUiContainer.style.zIndex"), "expected conflict overlay z-index to be driven by CSS");
  assert.ok(
    !text.includes("this.conflictUiContainer.style.pointerEvents"),
    "expected conflict overlay pointer-events to be driven by CSS",
  );

  // Toast/dialog roots: positioning + dialog chrome should be CSS-driven.
  assert.ok(!text.includes("toastRoot.style."), "expected conflict toast root to avoid static inline styles");
  assert.ok(!text.includes("dialogRoot.style."), "expected conflict dialog root to avoid static inline styles");
  assert.ok(
    !text.includes("structuralToastRoot.style."),
    "expected structural conflict toast root to avoid static inline styles",
  );
  assert.ok(
    !text.includes("structuralDialogRoot.style."),
    "expected structural conflict dialog root to avoid static inline styles",
  );

  // Sanity: ensure the expected CSS classes are applied.
  assert.match(conflictUiSection, /conflictUiContainer\.classList\.add\("conflict-ui-overlay"\)/);
  assert.match(conflictUiSection, /toastRoot\.classList\.add\("conflict-ui-toast-root"\)/);
  assert.match(conflictUiSection, /dialogRoot\.classList\.add\("conflict-ui-dialog-root"\)/);
  assert.match(conflictUiSection, /structuralToastRoot\.classList\.add\("structural-conflict-ui-toast-root"\)/);
  assert.match(conflictUiSection, /structuralDialogRoot\.classList\.add\("structural-conflict-ui-dialog-root"\)/);
});

test("conflict overlay classes are defined in conflicts.css", async () => {
  const cssPath = path.join(desktopRoot, "src/styles/conflicts.css");
  const css = await readFile(cssPath, "utf8");

  assert.match(css, /\.conflict-ui-overlay\s*\{/);
  assert.match(css, /position:\s*absolute\s*;/);
  assert.match(css, /pointer-events:\s*none\s*;/);

  assert.match(css, /\.conflict-ui-toast-root\s*\{/);
  assert.match(css, /\.structural-conflict-ui-toast-root\s*\{/);

  assert.match(css, /\.conflict-ui-dialog-root,\s*\n\.structural-conflict-ui-dialog-root\s*\{/);
  assert.match(css, /background:\s*var\(--dialog-bg\)\s*;/);
  assert.match(css, /border:\s*1px solid var\(--dialog-border\)\s*;/);
  assert.match(css, /box-shadow:\s*var\(--dialog-shadow\)\s*;/);
});

test("conflict styles are imported from main.ts", async () => {
  const mainPath = path.join(desktopRoot, "src/main.ts");
  const main = stripComments(await readFile(mainPath, "utf8"));
  assert.match(main, /^\s*import\s+["']\.\/styles\/conflicts\.css["']\s*;?/m);
});
