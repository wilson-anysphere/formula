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

  const reactMountSection = extractSection(source, "function renderReactPanel", "function renderDomPanel");
  assert.equal(
    /\.style\./.test(reactMountSection),
    false,
    "React panel mount container should not set inline styles; use a CSS class instead",
  );
  assert.match(
    reactMountSection,
    /dock-panel__mount/,
    "React panel mount container should apply the dock-panel__mount CSS class",
  );
  assert.match(
    reactMountSection,
    /panel-mount--fill-column/,
    "React panel mount container should apply the panel-mount--fill-column CSS class",
  );

  const domMountSection = extractSection(source, "function renderDomPanel", "function makeBodyFillAvailableHeight");
  assert.equal(
    /\.style\./.test(domMountSection),
    false,
    "DOM panel mount container should not set inline styles; use a CSS class instead",
  );
  assert.match(
    domMountSection,
    /dock-panel__mount/,
    "DOM panel mount container should apply the dock-panel__mount CSS class",
  );
  assert.match(
    domMountSection,
    /panel-mount--fill-column/,
    "DOM panel mount container should apply the panel-mount--fill-column CSS class",
  );

  const bodyFillSection = extractSection(source, "function makeBodyFillAvailableHeight", "function renderPanelBody");
  assert.equal(
    /\.style\./.test(bodyFillSection),
    false,
    "makeBodyFillAvailableHeight should not set inline styles; use a CSS class instead",
  );
  assert.match(
    bodyFillSection,
    /dock-panel__body--fill/,
    "makeBodyFillAvailableHeight should apply the dock-panel__body--fill CSS class",
  );
});

