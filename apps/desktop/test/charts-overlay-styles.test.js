import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");

test("shared-grid chart pane hosts use CSS classes (no static inline styles)", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  const start = text.indexOf("private initSharedChartPanes(): void {");
  assert.ok(start !== -1, "expected SpreadsheetApp.initSharedChartPanes() to exist");

  const end = text.indexOf("private syncSharedChartPanes", start);
  assert.ok(end !== -1, "expected SpreadsheetApp.syncSharedChartPanes() after initSharedChartPanes()");

  const block = text.slice(start, end);

  // Ensure pane hosts are class-based.
  assert.match(block, /chart-pane-top-left/);
  assert.match(block, /chart-pane--top-left/);
  assert.match(block, /chart-pane-top-right/);
  assert.match(block, /chart-pane--top-right/);
  assert.match(block, /chart-pane-bottom-left/);
  assert.match(block, /chart-pane--bottom-left/);
  assert.match(block, /chart-pane-bottom-right/);
  assert.match(block, /chart-pane--bottom-right/);

  // Ensure stable presentation styles are not set inline anymore.
  assert.ok(!block.includes("pane.style.position"), "expected pane host position to be driven by CSS");
  assert.ok(!block.includes("pane.style.pointerEvents"), "expected pane host pointer-events to be driven by CSS");
  assert.ok(!block.includes("pane.style.overflow"), "expected pane host overflow to be driven by CSS");
  assert.ok(!block.includes("pane.style.left"), "expected pane host default geometry to be driven by CSS");
  assert.ok(!block.includes("pane.style.top"), "expected pane host default geometry to be driven by CSS");
  assert.ok(!block.includes("pane.style.width"), "expected pane host default geometry to be driven by CSS");
  assert.ok(!block.includes("pane.style.height"), "expected pane host default geometry to be driven by CSS");
});

test("chart overlay hosts are styled via charts-overlay.css", async () => {
  const mainPath = path.join(desktopRoot, "src/main.ts");
  const main = await readFile(mainPath, "utf8");
  assert.match(main, /["']\.\/styles\/charts-overlay\.css["']/);

  const cssPath = path.join(desktopRoot, "src/styles/charts-overlay.css");
  const css = await readFile(cssPath, "utf8");

  assert.match(css, /\.chart-pane\s*\{/);
  assert.match(css, /position:\s*absolute\s*;/);
  assert.match(css, /pointer-events:\s*none\s*;/);
  assert.match(css, /overflow:\s*hidden\s*;/);

  // Ensure the chart overlay remains clipped under headers.
  assert.match(css, /\.chart-layer\s*\{/);
  assert.match(css, /overflow:\s*hidden\s*;/);

  // Ensure chart host divs can be created without inline presentation styles.
  assert.match(css, /\.chart-object\s*\{/);
});

test("shared-grid pane geometry stays dynamic (inline left/top/width/height/display)", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  const start = text.indexOf("private syncSharedChartPanes(viewport: GridViewportState): void {");
  assert.ok(start !== -1, "expected SpreadsheetApp.syncSharedChartPanes() to exist");

  const end = text.indexOf("private chartAnchorToViewportRect", start);
  assert.ok(end !== -1, "expected chartAnchorToViewportRect() after syncSharedChartPanes()");

  const block = text.slice(start, end);

  assert.match(block, /pane\.style\.left\s*=/);
  assert.match(block, /pane\.style\.top\s*=/);
  assert.match(block, /pane\.style\.width\s*=/);
  assert.match(block, /pane\.style\.height\s*=/);
  assert.match(block, /pane\.style\.display\s*=/);
});

test("shared-grid overlay stacking uses CSS classes (no zIndex inline styles)", async () => {
  const spreadsheetAppPath = path.join(desktopRoot, "src/app/spreadsheetApp.ts");
  const text = await readFile(spreadsheetAppPath, "utf8");

  assert.match(text, /chartLayer\.classList\.add\("chart-layer--shared"\)/);
  assert.match(text, /selectionCanvas\.classList\.add\("grid-canvas--shared-selection"\)/);
  assert.match(text, /outlineLayer\.classList\.add\("outline-layer--shared"\)/);

  // Sanity: we should not rely on these specific inline z-indices anymore.
  assert.ok(!text.includes('chartLayer.style.zIndex = "2"'));
  assert.ok(!text.includes('selectionCanvas.style.zIndex = "3"'));
  assert.ok(!text.includes('outlineLayer.style.zIndex = "4"'));
});
