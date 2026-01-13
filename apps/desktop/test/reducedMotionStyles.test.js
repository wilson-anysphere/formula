import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DESKTOP_ROOT = path.join(__dirname, "..");

function readDesktopFile(...segments) {
  return fs.readFileSync(path.join(DESKTOP_ROOT, ...segments), "utf8");
}

function extractTransitionDeclarations(css) {
  return [...css.matchAll(/transition\s*:\s*([\s\S]*?);/g)].map((match) => match[1]?.trim() ?? "");
}

function stripCssComments(css) {
  // CSS in this repo uses block comments only. Strip them so lint-style assertions
  // don't trip over explanatory text like "16ms" in comments.
  return css.replace(/\/\*[\s\S]*?\*\//g, "");
}

function collectCssFiles(dirPath) {
  /** @type {string[]} */
  const out = [];
  for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      out.push(...collectCssFiles(fullPath));
      continue;
    }
    if (entry.isFile() && entry.name.endsWith(".css")) out.push(fullPath);
  }
  return out;
}

test("Reduced motion overrides set motion tokens to 0ms", () => {
  const css = readDesktopFile("src", "styles", "tokens.css");

  assert.match(
    css,
    /:root\[data-reduced-motion="true"\][\s\S]*--motion-duration:\s*0ms;[\s\S]*--motion-duration-fast:\s*0ms;[\s\S]*--motion-ease:\s*linear;/,
    "Expected tokens.css to set --motion-duration/--motion-duration-fast to 0ms when data-reduced-motion=\"true\"",
  );

  assert.match(
    css,
    /@media\s*\(prefers-reduced-motion:\s*reduce\)[\s\S]*:root[\s\S]*--motion-duration:\s*0ms;[\s\S]*--motion-duration-fast:\s*0ms;[\s\S]*--motion-ease:\s*linear;/,
    "Expected tokens.css to set --motion-duration/--motion-duration-fast to 0ms under prefers-reduced-motion: reduce",
  );
});

test("Desktop CSS transitions use motion tokens (so reduced motion disables animation)", () => {
  const srcRoot = path.join(DESKTOP_ROOT, "src");
  const files = collectCssFiles(srcRoot);
  assert.ok(files.length > 0, "Expected to find CSS files under apps/desktop/src");

  for (const filePath of files) {
    const rawCss = fs.readFileSync(filePath, "utf8");
    const css = stripCssComments(rawCss);
    const relPath = path.relative(DESKTOP_ROOT, filePath);
    const isTokens = path.basename(filePath) === "tokens.css";

    assert.equal(
      /\banimation\s*:/.test(css) || /@keyframes\b/.test(css),
      false,
      `Unexpected CSS animations in ${relPath}; animations must be gated behind reduced-motion`,
    );

    for (const transition of extractTransitionDeclarations(css)) {
      if (!transition || transition.toLowerCase() === "none") continue;
      assert.match(
        transition,
        /var\(--motion-duration(?:-fast)?\)/,
        `Expected transition declarations in ${relPath} to use --motion-duration tokens`,
      );
      assert.match(
        transition,
        /var\(--motion-ease\)/,
        `Expected transition declarations in ${relPath} to use --motion-ease token`,
      );
    }

    // Guard against accidentally hardcoding durations in leaf CSS files (token definitions live in tokens.css).
    if (!isTokens) {
      assert.equal(
        /\b\d+(?:\.\d+)?(?:ms|s)\b/.test(css),
        false,
        `Expected ${relPath} to avoid hard-coded time values; use motion tokens instead`,
      );
    }
  }
});

test("Sheet tabs disable smooth scrolling under reduced motion", () => {
  const uiCss = readDesktopFile("src", "styles", "ui.css");

  assert.match(
    uiCss,
    /(?:html|:root)\[data-reduced-motion="true"\]\s+#sheet-tabs\.sheet-bar\s+\.sheet-tabs\s*\{[\s\S]*scroll-behavior:\s*auto;/,
    "Expected ui.css to disable smooth scrolling for sheet tabs when data-reduced-motion=\"true\"",
  );

  assert.match(
    uiCss,
    /@media\s*\(prefers-reduced-motion:\s*reduce\)[\s\S]*#sheet-tabs\.sheet-bar\s+\.sheet-tabs[\s\S]*scroll-behavior:\s*auto;/,
    "Expected ui.css to disable smooth scrolling for sheet tabs under prefers-reduced-motion: reduce",
  );

  const tabStripSrc = readDesktopFile("src", "sheets", "SheetTabStrip.tsx");
  assert.match(
    tabStripSrc,
    /data-reduced-motion[\s\S]*behavior:\s*reducedMotion\s*\?\s*"auto"\s*:\s*"smooth"/,
    "Expected SheetTabStrip to respect reduced-motion when choosing scroll behavior",
  );
});
