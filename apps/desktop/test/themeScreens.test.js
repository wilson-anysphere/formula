import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { expectSnapshot } from "./snapshot.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function parseVarsFromBlock(blockBody) {
  /** @type {Record<string, string>} */
  const vars = {};
  const regex = /--([a-z0-9-]+)\s*:\s*([^;]+);/gi;
  let match = null;
  while ((match = regex.exec(blockBody))) {
    vars[match[1]] = match[2].trim();
  }
  return vars;
}

function loadThemeVars(theme) {
  const tokensPath = path.join(__dirname, "..", "src", "styles", "tokens.css");
  const css = fs.readFileSync(tokensPath, "utf8");

  const rootMatch = css.match(/:root\s*\{([\s\S]*?)\}/);
  if (!rootMatch) throw new Error("tokens.css missing :root block");

  const base = parseVarsFromBlock(rootMatch[1]);

  if (theme === "light") return base;

  const themeMatch =
    theme === "dark"
      ? css.match(/:root\[data-theme="dark"\]\s*\{([\s\S]*?)\}/)
      : css.match(/:root\[data-theme="high-contrast"\]\s*\{([\s\S]*?)\}/);
  if (!themeMatch) throw new Error(`tokens.css missing ${theme} theme block`);

  return { ...base, ...parseVarsFromBlock(themeMatch[1]) };
}

function renderAppShell(theme) {
  const tokens = loadThemeVars(theme);
  const style = Object.entries(tokens)
    // Sort by code point order so snapshot ordering is stable across locales.
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .map(([name, value]) => `--${name}: ${value}`)
    .join("; ");

  return [
    `<!-- This snapshot intentionally inlines resolved CSS variables for ${theme}. -->`,
    `<div class="formula-app" data-theme="${theme}" style="${style}">`,
    `  <div class="formula-bar">`,
    `    <span class="formula-bar__label">fx</span>`,
    `    <input class="formula-bar__input" value="=SUM(A1:A10)" />`,
    `  </div>`,
    ``,
    `  <div class="grid">`,
    `    <div class="grid__headers">`,
    `      <div class="grid__header-cell"></div>`,
    `      <div class="grid__header-cell">A</div>`,
    `      <div class="grid__header-cell">B</div>`,
    `      <div class="grid__header-cell">C</div>`,
    `      <div class="grid__header-cell">D</div>`,
    `      <div class="grid__header-cell">E</div>`,
    `      <div class="grid__header-cell">F</div>`,
    `    </div>`,
    `    <div class="grid__body">`,
    `      <div class="grid__row-header">1</div>`,
    `      <div class="grid__cell">Product</div>`,
    `      <div class="grid__cell">Q1</div>`,
    `      <div class="grid__cell">Q2</div>`,
    `      <div class="grid__cell">Q3</div>`,
    `      <div class="grid__cell">Q4</div>`,
    `      <div class="grid__cell">Total</div>`,
    `      <div class="grid__row-header">2</div>`,
    `      <div class="grid__cell grid__cell--selected">Alpha</div>`,
    `      <div class="grid__cell">1,234</div>`,
    `      <div class="grid__cell">2,345</div>`,
    `      <div class="grid__cell">3,456</div>`,
    `      <div class="grid__cell">4,567</div>`,
    `      <div class="grid__cell">11,602</div>`,
    `    </div>`,
    `  </div>`,
    ``,
    `  <div class="panel">`,
    `    <h2 class="panel__title">AI Assistant</h2>`,
    `    <div>Ask a question about your data...</div>`,
    `  </div>`,
    ``,
    `  <div class="command-palette">`,
    `    <input class="command-palette__input" value="> Insert pivot table" />`,
    `    <ul class="command-palette__list">`,
    `      <li class="command-palette__item" aria-selected="true">Insert Pivot Table</li>`,
    `      <li class="command-palette__item">Insert Chart</li>`,
    `    </ul>`,
    `  </div>`,
    ``,
    `  <div class="dialog">`,
    `    <h3 class="dialog__title">Format Cells</h3>`,
    `    <div>Number, Alignment, Font…</div>`,
    `  </div>`,
    `</div>`,
    ``
  ].join("\n");
}

test("app shell snapshot (light)", () => {
  expectSnapshot("app-shell.light.html", renderAppShell("light"));
});

test("app shell snapshot (dark)", () => {
  expectSnapshot("app-shell.dark.html", renderAppShell("dark"));
});

test("app shell snapshot (high contrast)", () => {
  expectSnapshot("app-shell.high-contrast.html", renderAppShell("high-contrast"));
});
