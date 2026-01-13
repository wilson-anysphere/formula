import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("CellEditorOverlay avoids inline display/z-index style assignments", () => {
  const filePath = path.join(__dirname, "..", "src", "editor", "cellEditorOverlay.ts");
  const content = fs.readFileSync(filePath, "utf8");

  const forbiddenAssignments = [
    // Direct property assignments (e.g. `this.element.style.display = "none"`).
    /\.style\.display\s*=/,
    /\.style\s*\[\s*["']display["']\s*\]\s*=/,
    /\.style\.zIndex\s*=/,
    /\.style\s*\[\s*["']zIndex["']\s*\]\s*=/,
    /\.style\s*\[\s*["']z-index["']\s*\]\s*=/,
    // Alias assignments (e.g. `const style = this.element.style; style.display = "none"`).
    /\bstyle\s*\.display\s*=/,
    /\bstyle\s*\[\s*["']display["']\s*\]\s*=/,
    /\bstyle\s*\.zIndex\s*=/,
    /\bstyle\s*\[\s*["']zIndex["']\s*\]\s*=/,
    /\bstyle\s*\[\s*["']z-index["']\s*\]\s*=/,
    // setProperty (also mutates inline styles).
    /\.style\.setProperty\(\s*["']display["']\s*,/,
    /\bstyle\.setProperty\(\s*["']display["']\s*,/,
    /\.style\.setProperty\(\s*["']z-index["']\s*,/,
    /\bstyle\.setProperty\(\s*["']z-index["']\s*,/,
    /\.style\.setProperty\(\s*["']zIndex["']\s*,/,
    /\bstyle\.setProperty\(\s*["']zIndex["']\s*,/,
    // removeProperty (also mutates inline styles).
    /\.style\.removeProperty\(\s*["']display["']\s*\)/,
    /\bstyle\.removeProperty\(\s*["']display["']\s*\)/,
    /\.style\.removeProperty\(\s*["']z-index["']\s*\)/,
    /\bstyle\.removeProperty\(\s*["']z-index["']\s*\)/,
    /\.style\.removeProperty\(\s*["']zIndex["']\s*\)/,
    /\bstyle\.removeProperty\(\s*["']zIndex["']\s*\)/,
    // Object.assign(...style, { display/zIndex: ... })
    /Object\.assign\(\s*[^,]+\.style\s*,[\s\S]*?\bdisplay\s*:/,
    /Object\.assign\(\s*[^,]+\.style\s*,[\s\S]*?\bzIndex\s*:/,
    /Object\.assign\(\s*style\s*,[\s\S]*?\bdisplay\s*:/,
    /Object\.assign\(\s*style\s*,[\s\S]*?\bzIndex\s*:/,
    // cssText assignments that include forbidden properties.
    /\.style\.cssText\s*\+?=[\s\S]{0,200}\bdisplay\s*:/,
    /\.style\.cssText\s*\+?=[\s\S]{0,200}\bz-index\s*:/,
    /\bstyle\s*\.cssText\s*\+?=[\s\S]{0,200}\bdisplay\s*:/,
    /\bstyle\s*\.cssText\s*\+?=[\s\S]{0,200}\bz-index\s*:/,
    // setAttribute("style", "...display..." / "...z-index...")
    /setAttribute\(\s*["']style["']\s*,[\s\S]{0,200}\bdisplay\s*:/,
    /setAttribute\(\s*["']style["']\s*,[\s\S]{0,200}\bz-index\s*:/,
  ];

  for (const pattern of forbiddenAssignments) {
    assert.equal(
      pattern.test(content),
      false,
      `CellEditorOverlay should not assign inline styles for display/zIndex (matched ${pattern})`,
    );
  }
});

test("CellEditorOverlay static visibility + z-index are defined in CSS", () => {
  const cssPath = path.join(__dirname, "..", "src", "styles", "shell.css");
  const css = fs.readFileSync(cssPath, "utf8");

  const baseRule = css.match(/\.cell-editor\s*\{([\s\S]*?)\}/);
  assert.ok(baseRule, "Expected shell.css to define a .cell-editor rule");
  const baseBody = baseRule[1] ?? "";

  // Stacking must be handled in CSS (especially for shared-grid mode).
  const zIndexMatch = baseBody.match(/\bz-index\s*:\s*([^;]+);?/);
  assert.ok(zIndexMatch, "Expected .cell-editor styles to set z-index via CSS");
  const zIndexValue = (zIndexMatch?.[1] ?? "").trim();
  if (/^\d+$/.test(zIndexValue)) {
    const parsed = Number.parseInt(zIndexValue, 10);
    const chartsCssPath = path.join(__dirname, "..", "src", "styles", "charts-overlay.css");
    const chartsCss = fs.readFileSync(chartsCssPath, "utf8");
    // Shared-grid overlay layers are explicitly z-indexed in charts-overlay.css. Keep the
    // editor above all of them so it can sit on top of selection/charts overlays.
    const sharedSelectors = [
      /\.grid-canvas--shared-chart\s*\{[\s\S]*?\bz-index\s*:\s*(\d+)/,
      /\.grid-canvas--shared-selection\s*\{[\s\S]*?\bz-index\s*:\s*(\d+)/,
      /\.outline-layer--shared\s*\{[\s\S]*?\bz-index\s*:\s*(\d+)/,
    ];
    const sharedZ = sharedSelectors
      .map((re) => chartsCss.match(re)?.[1])
      .filter((v) => typeof v === "string" && v.trim() !== "")
      .map((v) => Number.parseInt(String(v), 10))
      .filter((n) => Number.isFinite(n));
    const maxShared = sharedZ.length === 0 ? 0 : Math.max(...sharedZ);
    assert.ok(
      parsed > maxShared,
      `Expected .cell-editor z-index (${zIndexValue}) to stay above shared-grid overlay layers (max ${maxShared})`,
    );
  }

  const baseHidden = /\bdisplay\s*:\s*none\s*;?/.test(baseBody);
  const hiddenRule = css.match(/\.cell-editor[^{]*\[\s*hidden\s*\][^{]*\{([\s\S]*?)\}/);
  const hiddenViaAttr = hiddenRule ? /\bdisplay\s*:\s*none\s*;?/.test(hiddenRule[1] ?? "") : false;

  assert.ok(
    baseHidden || hiddenViaAttr,
    "Expected the cell editor to be hidden by default via CSS (display:none or [hidden])",
  );

  // If the base rule hides the editor, ensure we have a modifier selector that
  // makes it visible again.
  if (baseHidden) {
    const openRule = css.match(/\.cell-editor[^{]*cell-editor--open[^{]*\{([\s\S]*?)\}/);
    assert.ok(openRule, "Expected shell.css to define a .cell-editor--open modifier rule");
    assert.ok(
      /\bdisplay\s*:\s*(?!none\b)[a-z-]+\s*;?/i.test(openRule[1] ?? ""),
      "Expected .cell-editor--open to set display to a visible value (not none)",
    );
  }
});

test("CellEditorOverlay keeps geometry dynamic via inline styles", () => {
  const filePath = path.join(__dirname, "..", "src", "editor", "cellEditorOverlay.ts");
  const content = fs.readFileSync(filePath, "utf8");

  for (const prop of ["left", "top", "width", "height"]) {
    assert.match(
      content,
      new RegExp(`\\.style\\.${prop}\\s*=`),
      `Expected CellEditorOverlay to keep ${prop} dynamic via inline style assignments`,
    );
  }
});

test("CellEditorOverlay only uses inline styles for dynamic geometry", () => {
  const filePath = path.join(__dirname, "..", "src", "editor", "cellEditorOverlay.ts");
  const content = fs.readFileSync(filePath, "utf8");

  const allowed = new Set(["left", "top", "width", "height"]);

  const dotMatches = [...content.matchAll(/\.style\.([a-zA-Z]+)\s*=/g)];
  const bracketMatches = [...content.matchAll(/\.style\s*\[\s*["']([a-zA-Z-]+)["']\s*\]\s*=/g)];

  const disallowed = [
    ...dotMatches.filter((m) => !allowed.has(m[1])),
    ...bracketMatches.filter((m) => !allowed.has(m[1])),
  ];

  assert.deepEqual(
    disallowed.map((m) => m[0]),
    [],
    "CellEditorOverlay should only set inline styles for left/top/width/height; move other presentation styles to CSS",
  );
});
