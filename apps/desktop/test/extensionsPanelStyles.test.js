import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "./sourceTextUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const desktopRoot = path.resolve(__dirname, "..");

test("ExtensionsPanel uses CSS classes (no React inline style props)", async () => {
  const panelPath = path.join(desktopRoot, "src", "extensions", "ExtensionsPanel.tsx");
  const panelSource = stripComments(await readFile(panelPath, "utf8"));

  assert.ok(
    !panelSource.includes("style={{"),
    "ExtensionsPanel.tsx should not use React inline style props; use token-driven CSS classes instead",
  );
  assert.ok(
    !panelSource.includes(".style."),
    "ExtensionsPanel.tsx should not assign element.style; use token-driven CSS classes instead",
  );
  assert.match(panelSource, /className\s*=\s*["'][^"']*extensions-panel\b/);

  const bodyPath = path.join(desktopRoot, "src", "extensions", "ExtensionPanelBody.tsx");
  const bodySource = stripComments(await readFile(bodyPath, "utf8"));

  assert.ok(
    !bodySource.includes("style={{"),
    "ExtensionPanelBody.tsx should not use React inline style props; use token-driven CSS classes instead",
  );
  assert.ok(
    !bodySource.includes(".style."),
    "ExtensionPanelBody.tsx should not assign element.style; use token-driven CSS classes instead",
  );
  assert.match(bodySource, /className\s*=\s*["'][^"']*extension-webview\b/);
});

test("extensions.css is imported and defines required classes", async () => {
  const mainPath = path.join(desktopRoot, "src", "main.ts");
  const main = stripComments(await readFile(mainPath, "utf8"));
  assert.match(main, /^\s*import\s+["']\.\/styles\/extensions\.css["']\s*;?/m);

  const cssPath = path.join(desktopRoot, "src", "styles", "extensions.css");
  const css = await readFile(cssPath, "utf8");
  assert.match(css, /\.extensions-panel\s*\{/);
  assert.match(css, /\.extension-webview\s*\{/);
});
