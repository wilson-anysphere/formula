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

test("panelBodyRenderer.tsx avoids inline styles (class-driven panel mounts)", () => {
  const filePath = path.join(__dirname, "..", "src", "panels", "panelBodyRenderer.tsx");
  const source = fs.readFileSync(filePath, "utf8");
  const versionHistoryPath = path.join(
    __dirname,
    "..",
    "src",
    "panels",
    "version-history",
    "CollabVersionHistoryPanel.tsx",
  );
  const versionHistorySource = fs.readFileSync(versionHistoryPath, "utf8");
  const branchManagerPath = path.join(
    __dirname,
    "..",
    "src",
    "panels",
    "branch-manager",
    "CollabBranchManagerPanel.tsx",
  );
  const branchManagerSource = fs.readFileSync(branchManagerPath, "utf8");

  assert.equal(
    /<[^>]*\bstyle\s*=\s*\{/.test(source),
    false,
    "panelBodyRenderer.tsx should avoid React inline styles (style={...}); use CSS classes instead",
  );
  assert.equal(
    /<[^>]*\bstyle\s*=\s*["']/i.test(source),
    false,
    "panelBodyRenderer.tsx should avoid HTML inline styles (style=\"...\"); use CSS classes instead",
  );
  assert.equal(
    /\.style\./.test(source),
    false,
    "panelBodyRenderer.tsx should not assign DOM inline styles via `.style.*`; use CSS classes instead",
  );
  assert.equal(
    /\[\s*["']style["']\s*\]/.test(source),
    false,
    "panelBodyRenderer.tsx should not access the style attribute via bracket notation; use CSS classes instead",
  );
  assert.equal(
    /setAttribute\(\s*["']style["']/.test(source) || /setAttributeNS\(\s*[^,]+,\s*["']style["']/.test(source),
    false,
    "panelBodyRenderer.tsx should not set style attributes via setAttribute; use CSS classes instead",
  );

  for (const forbidden of [
    'container.style.height = "100%"',
    'container.style.display = "flex"',
    'body.style.flex = "1"',
  ]) {
    assert.equal(
      source.includes(forbidden),
      false,
      `panelBodyRenderer.tsx should not contain static inline style assignment: ${forbidden}`,
    );
  }

  for (const legacyClass of ["dock-panel__mount", "panel-mount--fill-column", "dock-panel__body--fill"]) {
    assert.equal(
      source.includes(legacyClass),
      false,
      `panelBodyRenderer.tsx should not reference legacy CSS class ${legacyClass}; prefer panel-body__container/panel-body--fill`,
    );
  }

  assert.equal(
    /\bstyle=\{\{/.test(versionHistorySource),
    false,
    "CollabVersionHistoryPanel should not use React inline styles; use CSS classes in workspace.css instead",
  );
  assert.equal(
    /\.style\./.test(versionHistorySource),
    false,
    "CollabVersionHistoryPanel should not assign DOM inline styles via `.style.*`; use CSS classes instead",
  );

  assert.equal(
    /\bstyle=\{\{/.test(branchManagerSource),
    false,
    "CollabBranchManagerPanel should not use React inline styles; use CSS classes in workspace.css instead",
  );

  const reactMountSection = extractSection(source, "function renderReactPanel", "function renderDomPanel");
  assert.equal(
    /\.style\./.test(reactMountSection),
    false,
    "React panel mount container should not set inline styles; use a CSS class instead",
  );
  for (const className of ["panel-body__container"]) {
    assert.match(reactMountSection, new RegExp(className), `React panel mount container should apply the ${className} CSS class`);
  }

  const domMountSection = extractSection(source, "function renderDomPanel", "function makeBodyFillAvailableHeight");
  assert.equal(
    /\.style\./.test(domMountSection),
    false,
    "DOM panel mount container should not set inline styles; use a CSS class instead",
  );
  for (const className of ["panel-body__container"]) {
    assert.match(domMountSection, new RegExp(className), `DOM panel mount container should apply the ${className} CSS class`);
  }

  const bodyFillSection = extractSection(source, "function makeBodyFillAvailableHeight", "function renderPanelBody");
  assert.equal(
    /\.style\./.test(bodyFillSection),
    false,
    "makeBodyFillAvailableHeight should not set inline styles; use a CSS class instead",
  );
  assert.match(bodyFillSection, /panel-body--fill/, "makeBodyFillAvailableHeight should apply the panel-body--fill CSS class");

  const renderPanelBodySection = extractSection(source, "function renderPanelBody", "function cleanup");
  assert.match(
    renderPanelBodySection,
    /classList\.remove\([^)]*panel-body--fill/,
    "renderPanelBody should clear panel-body--fill before applying the next panel's layout",
  );
});
