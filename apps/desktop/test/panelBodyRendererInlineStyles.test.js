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

test("panelBodyRenderer.tsx avoids inline style assignments for dock panel mounts", () => {
  const filePath = path.join(__dirname, "..", "src", "panels", "panelBodyRenderer.tsx");
  const source = fs.readFileSync(filePath, "utf8");

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

  const versionHistorySection = extractSection(source, "function CollabVersionHistoryPanel", "function CollabBranchManagerPanel");
  assert.equal(
    /\bstyle=\{\{/.test(versionHistorySection),
    false,
    "CollabVersionHistoryPanel should not use React inline styles; use CSS classes in workspace.css instead",
  );

  const branchManagerSection = extractSection(source, "function CollabBranchManagerPanel", "export interface PanelBodyRendererOptions");
  assert.equal(
    /\bstyle=\{\{/.test(branchManagerSection),
    false,
    "CollabBranchManagerPanel should not use React inline styles; use CSS classes in workspace.css instead",
  );

  const reactMountSection = extractSection(source, "function renderReactPanel", "function renderDomPanel");
  assert.equal(
    /\.style\./.test(reactMountSection),
    false,
    "React panel mount container should not set inline styles; use a CSS class instead",
  );
  for (const className of ["dock-panel__mount", "panel-mount--fill-column", "panel-body__container"]) {
    assert.match(reactMountSection, new RegExp(className), `React panel mount container should apply the ${className} CSS class`);
  }

  const domMountSection = extractSection(source, "function renderDomPanel", "function makeBodyFillAvailableHeight");
  assert.equal(
    /\.style\./.test(domMountSection),
    false,
    "DOM panel mount container should not set inline styles; use a CSS class instead",
  );
  for (const className of ["dock-panel__mount", "panel-mount--fill-column", "panel-body__container"]) {
    assert.match(domMountSection, new RegExp(className), `DOM panel mount container should apply the ${className} CSS class`);
  }

  const bodyFillSection = extractSection(source, "function makeBodyFillAvailableHeight", "function renderPanelBody");
  assert.equal(
    /\.style\./.test(bodyFillSection),
    false,
    "makeBodyFillAvailableHeight should not set inline styles; use a CSS class instead",
  );
  for (const className of ["dock-panel__body--fill", "panel-body--fill"]) {
    assert.match(bodyFillSection, new RegExp(className), `makeBodyFillAvailableHeight should apply the ${className} CSS class`);
  }
});
