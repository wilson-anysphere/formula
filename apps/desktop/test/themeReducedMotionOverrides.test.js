import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("ui.css disables smooth scrolling when reduced motion is enabled", () => {
  const uiCssPath = path.join(__dirname, "..", "src", "styles", "ui.css");
  const css = fs.readFileSync(uiCssPath, "utf8");

  // Guardrail: keep the smooth-scroll baseline for users without reduced motion enabled.
  assert.match(
    css,
    /#sheet-tabs\.sheet-bar\s+\.sheet-tabs\s*\{[\s\S]*?scroll-behavior:\s*smooth\s*;/,
  );

  // Runtime reduced-motion flag (data attribute on <html>).
  assert.match(
    css,
    /html\[data-reduced-motion=["']true["']\]\s+#sheet-tabs\.sheet-bar\s+\.sheet-tabs\s*\{[^}]*scroll-behavior:\s*auto\s*;/,
  );

  // OS-level reduced-motion preference.
  const mediaIdx = css.indexOf("@media (prefers-reduced-motion: reduce)");
  assert.ok(mediaIdx >= 0, "Expected ui.css to include a prefers-reduced-motion: reduce media query");
  assert.match(
    css.slice(mediaIdx),
    /#sheet-tabs\.sheet-bar\s+\.sheet-tabs\s*\{[^}]*scroll-behavior:\s*auto\s*;/,
  );
});

test("SheetTabStrip avoids smooth scrollBy when reduced motion is enabled", () => {
  const sheetTabStripPath = path.join(__dirname, "..", "src", "sheets", "SheetTabStrip.tsx");
  const src = fs.readFileSync(sheetTabStripPath, "utf8");

  // Guardrail: don't hardcode smooth scrolling in JS without a reduced-motion escape hatch.
  assert.doesNotMatch(
    src,
    /scrollBy\(\s*\{\s*left:\s*delta\s*,\s*behavior:\s*["']smooth["']\s*\}\s*\)/,
  );

  // Expect a conditional that can fall back to auto scrolling.
  assert.match(
    src,
    /scrollBy\(\s*\{\s*left:\s*delta\s*,\s*behavior:\s*[^}]*\?\s*["']auto["']\s*:\s*["']smooth["'][^}]*\}\s*\)/,
  );

  // Ensure we respect both runtime + OS reduced motion signals.
  assert.match(src, /data-reduced-motion/);
  assert.match(src, /prefers-reduced-motion:\s*reduce/);
});
