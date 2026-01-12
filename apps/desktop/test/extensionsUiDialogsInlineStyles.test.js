import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function extractSection(source, startMarker, endMarker) {
  const startIdx = source.indexOf(startMarker);
  assert.ok(startIdx !== -1, `Expected to find start marker: ${startMarker}`);

  if (!endMarker) return source.slice(startIdx);

  const endIdx = source.indexOf(endMarker, startIdx);
  assert.ok(endIdx !== -1, `Expected to find end marker: ${endMarker}`);

  return source.slice(startIdx, endIdx);
}

test("extension UI dialogs avoid inline style assignments", () => {
  const uiPath = path.join(__dirname, "..", "src", "extensions", "ui.ts");
  const uiSource = fs.readFileSync(uiPath, "utf8");

  const inputBoxSection = extractSection(
    uiSource,
    "export async function showInputBox",
    "type QuickPickItem",
  );
  assert.equal(
    /\.style\b/.test(inputBoxSection),
    false,
    "showInputBox should not set inline styles; use token-based CSS classes instead",
  );
  assert.equal(
    /cssText\b/.test(inputBoxSection),
    false,
    "showInputBox should not assign cssText; use token-based CSS classes instead",
  );
  assert.equal(
    /setAttribute\(\s*["']style["']/.test(inputBoxSection),
    false,
    "showInputBox should not set style attributes; use token-based CSS classes instead",
  );
  assert.match(
    inputBoxSection,
    /input\.className\s*=\s*"dialog__field"/,
    "showInputBox should apply the dialog__field CSS class",
  );
  assert.match(
    inputBoxSection,
    /controls\.className\s*=\s*"dialog__controls"/,
    "showInputBox should apply the dialog__controls CSS class",
  );
  assert.match(
    inputBoxSection,
    /dialog\.className\s*=\s*"dialog extensions-ui"/,
    "showInputBox should add an extensions-ui class to scope dialog-specific styling",
  );
  assert.match(
    inputBoxSection,
    /dialog\.setAttribute\(\s*["']aria-labelledby["']/,
    "showInputBox should associate an accessible name via aria-labelledby",
  );

  const quickPickSection = extractSection(uiSource, "export async function showQuickPick");
  assert.equal(
    /\.style\b/.test(quickPickSection),
    false,
    "showQuickPick should not set inline styles; use token-based CSS classes instead",
  );
  assert.equal(
    /cssText\b/.test(quickPickSection),
    false,
    "showQuickPick should not assign cssText; use token-based CSS classes instead",
  );
  assert.equal(
    /setAttribute\(\s*["']style["']/.test(quickPickSection),
    false,
    "showQuickPick should not set style attributes; use token-based CSS classes instead",
  );
  assert.match(
    quickPickSection,
    /list\.className\s*=\s*"quick-pick__list"/,
    "showQuickPick should apply the quick-pick__list CSS class",
  );
  assert.match(
    quickPickSection,
    /btn\.className\s*=\s*"quick-pick__item"/,
    "showQuickPick should apply the quick-pick__item CSS class",
  );
  assert.match(
    quickPickSection,
    /label\.className\s*=\s*"quick-pick__label"/,
    "showQuickPick should apply the quick-pick__label CSS class",
  );
  assert.match(
    quickPickSection,
    /sub\.className\s*=\s*"quick-pick__subtext"/,
    "showQuickPick should apply the quick-pick__subtext CSS class",
  );
  assert.match(
    quickPickSection,
    /dialog\.className\s*=\s*"dialog extensions-ui"/,
    "showQuickPick should add an extensions-ui class to scope dialog-specific styling",
  );
  assert.match(
    quickPickSection,
    /dialog\.setAttribute\(\s*["']aria-labelledby["']/,
    "showQuickPick should associate an accessible name via aria-labelledby",
  );

  const mainPath = path.join(__dirname, "..", "src", "main.ts");
  const mainSource = fs.readFileSync(mainPath, "utf8");
  assert.match(
    mainSource,
    /import\s+["']\.\/styles\/dialogs\.css["'];/,
    "main.ts should import dialogs.css so extension dialogs render correctly",
  );
  assert.match(
    mainSource,
    /import\s+["']\.\/styles\/extensions-ui\.css["'];/,
    "main.ts should import extensions-ui.css so extension dialogs render correctly",
  );
});
