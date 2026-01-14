import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");

test("SpreadsheetApp overlay canvases use CSS classes (no static inline styles)", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  const start = text.indexOf('this.gridCanvas = document.createElement("canvas");');
  assert.ok(start !== -1, "expected SpreadsheetApp to create overlay canvases via document.createElement");

  const end = text.indexOf("this.root.appendChild(this.outlineLayer);", start);
  assert.ok(end !== -1, "expected SpreadsheetApp to append outlineLayer after creating overlay canvases");

  const block = text.slice(start, end);

  // Ensure overlay layers are class-based.
  assert.match(block, /gridCanvas\.className\s*=\s*"grid-canvas grid-canvas--base"/);
  assert.match(block, /drawingCanvas\.className\s*=\s*"drawing-layer drawing-layer--overlay[^"]*grid-canvas--drawings"/);
  assert.match(block, /chartCanvas\.className\s*=\s*"grid-canvas grid-canvas--chart"/);
  assert.match(block, /selectionCanvas\.className\s*=\s*"grid-canvas grid-canvas--selection"/);
  assert.match(block, /chartSelectionCanvas\.className\s*=\s*"grid-canvas[^"]*chart-selection-canvas"/);

  // Shared-grid overlay stacking is expressed via CSS classes (see charts-overlay.css).
  assert.match(block, /drawingCanvas\.classList\.add\([\s\S]*?"drawing-layer--shared"[\s\S]*?"grid-canvas--shared-drawings"[\s\S]*?\)/);
  assert.match(block, /chartCanvas\.classList\.add\("grid-canvas--shared-chart"\)/);
  assert.match(block, /selectionCanvas\.classList\.add\("grid-canvas--shared-selection"\)/);
  assert.match(block, /chartSelectionCanvas\.classList\.add\("grid-canvas--shared-selection"\)/);
  assert.match(block, /outlineLayer\.classList\.add\("outline-layer--shared"\)/);

  // Ensure stable presentation styles are not set inline at mount time.
  assert.ok(!block.includes(".style.position"), "expected overlay positioning to be driven by CSS");
  // Some overlay canvases (e.g. chartSelectionCanvas) may explicitly disable pointer events inline.
  // Guard against accidentally wiring pointer-events inline on the primary overlay layers.
  assert.ok(
    !block.includes("drawingCanvas.style.pointerEvents"),
    "expected drawingCanvas pointer-events to be driven by CSS",
  );
  assert.ok(!block.includes("chartCanvas.style.pointerEvents"), "expected chartCanvas pointer-events to be driven by CSS");
  assert.ok(
    !block.includes("selectionCanvas.style.pointerEvents"),
    "expected selectionCanvas pointer-events to be driven by CSS",
  );
  assert.ok(!block.includes(".style.overflow"), "expected overlay overflow clipping to be driven by CSS");

  // Presence overlays share the selection z-index (charts-overlay.css) and rely on DOM insertion
  // order for tie-breaking (selection should remain on top of presence highlights).
  const presenceAppendIndex = block.search(/\bappendChild\(this\.presenceCanvas\)/);
  const selectionAppendIndex = block.search(/\bappendChild\(this\.selectionCanvas\)/);
  assert.ok(presenceAppendIndex !== -1, "expected SpreadsheetApp to append presenceCanvas when present");
  assert.ok(selectionAppendIndex !== -1, "expected SpreadsheetApp to append selectionCanvas");
  assert.ok(
    presenceAppendIndex < selectionAppendIndex,
    "expected presenceCanvas to be appended before selectionCanvas so selection renders above presence",
  );
});

test("chart + drawing overlay hosts are styled via charts-overlay.css", async () => {
  const mainPath = path.join(desktopRoot, "src/main.ts");
  const main = await readFile(mainPath, "utf8");
  assert.match(main, /["']\.\/styles\/charts-overlay\.css["']/);

  const cssPath = path.join(desktopRoot, "src/styles/charts-overlay.css");
  const css = await readFile(cssPath, "utf8");

  // Charts are rendered on a dedicated canvas layer.
  assert.match(css, /\.grid-canvas--chart\s*\{/);
  assert.match(css, /pointer-events:\s*none\s*;/);

  // Chart selection handles are rendered on a separate overlay canvas.
  assert.match(css, /\.chart-selection-canvas\s*\{/);
  assert.match(css, /\.chart-selection-canvas\s*\{[^}]*pointer-events:\s*none\s*;/);

  // Drawings (shapes/images) are rendered to a canvas clipped under headers.
  assert.match(css, /\.drawing-layer\s*\{/);
  assert.match(css, /position:\s*absolute\s*;/);
  assert.match(css, /pointer-events:\s*none\s*;/);
  assert.match(css, /overflow:\s*hidden\s*;/);

  // Shared grid drawings overlay canvases (split view) are also non-interactive.
  assert.match(css, /\.grid-canvas--drawings\s*\{[^}]*pointer-events:\s*none\s*;/);
  // Back-compat: allow overriding the older `--grid-z-drawing-overlay` token.
  assert.match(
    css,
    /\.grid-canvas--drawings\s*\{[\s\S]*?z-index:\s*var\(--grid-z-drawings-overlay,\s*var\(--grid-z-drawing-overlay,\s*3\)\)\s*;/,
  );

  // Collaboration presence overlay is non-interactive.
  assert.match(css, /\.grid-canvas--presence\s*\{[^}]*pointer-events:\s*none\s*;/);
  // Presence should participate in the selection overlay stack so it stays above charts/drawings.
  assert.match(
    css,
    /\.grid-canvas--presence\s*\{[\s\S]*?z-index:\s*var\(--grid-z-selection-overlay,\s*4\)\s*;/,
  );

  // Shared-grid overlay stacking is driven via CSS variables + semantic classes.
  assert.match(css, /--grid-z-chart-overlay\s*:/);
  assert.match(css, /--grid-z-drawings-overlay\s*:/);
  // Back-compat alias (older selectors used singular `drawing`).
  assert.match(css, /--grid-z-drawing-overlay\s*:/);
  assert.match(css, /--grid-z-selection-overlay\s*:/);
  assert.match(css, /--grid-z-outline-overlay\s*:/);

  // Presence canvas participates in the same overlay stack so it stays above charts/drawings.
  assert.match(css, /\.grid-canvas--presence\s*\{/);

  assert.match(css, /\.drawing-layer--shared\s*\{/);
  assert.match(css, /\.grid-canvas--shared-drawings\s*\{/);
  assert.match(css, /\.grid-canvas--shared-chart\s*\{/);
  assert.match(css, /\.grid-canvas--shared-selection\s*\{/);
  assert.match(css, /\.outline-layer--shared\s*\{/);

  // The selection canvas should override CanvasGridRenderer's inline z-index in shared-grid mode.
  assert.match(css, /\.grid-canvas--shared-selection\s*\{[\s\S]*?!important/);
});

test("chart + drawing overlay geometry stays dynamic (inline offsets/sizing)", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  // Drawings overlay is offset under headers.
  assert.match(text, /drawingCanvas\.style\.left\s*=/);
  assert.match(text, /drawingCanvas\.style\.top\s*=/);

  // Chart canvas is resized dynamically for DPR.
  assert.match(text, /chartCanvas\.style\.width\s*=/);
  assert.match(text, /chartCanvas\.style\.height\s*=/);
});

test("shared-grid overlay stacking uses CSS classes (no zIndex inline styles)", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  assert.match(text, /drawingCanvas\.classList\.add\([\s\S]*?"drawing-layer--shared"[\s\S]*?"grid-canvas--shared-drawings"[\s\S]*?\)/);
  assert.match(text, /chartCanvas\.classList\.add\("grid-canvas--shared-chart"\)/);
  assert.match(text, /selectionCanvas\.classList\.add\("grid-canvas--shared-selection"\)/);
  assert.match(text, /chartSelectionCanvas\.classList\.add\("grid-canvas--shared-selection"\)/);
  assert.match(text, /outlineLayer\.classList\.add\("outline-layer--shared"\)/);

  // Sanity: overlay stacking should be driven by CSS, not inline zIndex assignments.
  assert.ok(!text.includes("drawingCanvas.style.zIndex"), "expected no inline zIndex for drawingCanvas");
  assert.ok(!text.includes("chartCanvas.style.zIndex"), "expected no inline zIndex for chartCanvas");
  assert.ok(!text.includes("selectionCanvas.style.zIndex"), "expected no inline zIndex for selectionCanvas");
  assert.ok(!text.includes("chartSelectionCanvas.style.zIndex"), "expected no inline zIndex for chartSelectionCanvas");
  assert.ok(!text.includes("outlineLayer.style.zIndex"), "expected no inline zIndex for outlineLayer");
});

test("SpreadsheetApp assigns semantic layer classes to grid canvases + overlays", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  // Canvases/layers should be tagged with role-ish classes so CSS can target them
  // without relying on DOM insertion order.
  assert.match(text, /gridCanvas\.className\s*=\s*"grid-canvas grid-canvas--base"/);
  assert.match(text, /drawingCanvas\.className\s*=\s*"drawing-layer drawing-layer--overlay[^"]*grid-canvas--drawings"/);
  assert.match(text, /chartCanvas\.className\s*=\s*"grid-canvas grid-canvas--chart"/);
  assert.match(text, /chartSelectionCanvas\.className\s*=\s*"grid-canvas[^"]*chart-selection-canvas"/);
  assert.match(text, /referenceCanvas\.className\s*=\s*"grid-canvas grid-canvas--content"/);
  assert.match(text, /auditingCanvas\.className\s*=\s*"grid-canvas grid-canvas--auditing"/);
  assert.match(text, /selectionCanvas\.className\s*=\s*"grid-canvas grid-canvas--selection"/);
  assert.match(text, /presenceCanvas\.className\s*=\s*"grid-canvas grid-canvas--presence"/);
});

test("SecondaryGridView assigns semantic layer classes to shared-grid canvases", async () => {
  const secondaryGridPath = path.join(desktopRoot, "src/grid/splitView/secondaryGridView.ts");
  const text = await readFile(secondaryGridPath, "utf8");

  assert.match(text, /gridCanvas\.className\s*=\s*"grid-canvas grid-canvas--base"/);
  assert.match(text, /contentCanvas\.className\s*=\s*"grid-canvas grid-canvas--content"/);
  assert.match(text, /selectionCanvas\.className\s*=\s*"grid-canvas grid-canvas--selection"/);
});
