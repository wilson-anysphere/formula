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

function stripJsComments(src) {
  // Conservative comment stripping for lint-style assertions.
  // This is not a full parser, but it's good enough to avoid false positives from comments.
  return (
    src
      // Block comments
      .replace(/\/\*[\s\S]*?\*\//g, "")
      // Line comments
      .replace(/(^|[^:])\/\/.*$/gm, "$1")
  );
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

function collectSourceFiles(dirPath) {
  /** @type {string[]} */
  const out = [];
  for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
    const fullPath = path.join(dirPath, entry.name);
    if (entry.isDirectory()) {
      out.push(...collectSourceFiles(fullPath));
      continue;
    }
    if (!entry.isFile()) continue;
    if (!/\.(?:ts|tsx|js|jsx)$/.test(entry.name)) continue;
    out.push(fullPath);
  }
  return out;
}

test("Reduced motion overrides set motion tokens to 0ms", () => {
  const css = readDesktopFile("src", "styles", "tokens.css");

  // Token-based transitions: used throughout the UI so reduced motion can zero them out.
  assert.match(
    css,
    /:root\[data-reduced-motion\s*=\s*(?:"true"|'true'|true)\][\s\S]*--motion-duration:\s*0ms;[\s\S]*--motion-duration-fast:\s*0ms;[\s\S]*--motion-ease:\s*linear;/,
    "Expected tokens.css to set --motion-duration/--motion-duration-fast to 0ms when data-reduced-motion=\"true\"",
  );

  // Smooth scrolling is also motion-heavy; ensure reduced motion forces it off globally.
  assert.match(
    css,
    /:root\[data-reduced-motion\s*=\s*(?:"true"|'true'|true)\][\s\S]*scroll-behavior:\s*auto\s*!important\s*;/,
    "Expected tokens.css to disable smooth scrolling (scroll-behavior: auto) when data-reduced-motion=\"true\"",
  );

  assert.match(
    css,
    /:root\[data-reduced-motion\s*=\s*(?:"true"|'true'|true)\]\s*\*\s*\{[\s\S]*scroll-behavior:\s*auto\s*!important\s*;/,
    "Expected tokens.css to disable smooth scrolling for all descendants when data-reduced-motion=\"true\"",
  );

  assert.match(
    css,
    /@media\s*\(prefers-reduced-motion:\s*reduce\)[\s\S]*:root[\s\S]*--motion-duration:\s*0ms;[\s\S]*--motion-duration-fast:\s*0ms;[\s\S]*--motion-ease:\s*linear;/,
    "Expected tokens.css to set --motion-duration/--motion-duration-fast to 0ms under prefers-reduced-motion: reduce",
  );

  assert.match(
    css,
    /@media\s*\(prefers-reduced-motion:\s*reduce\)[\s\S]*:root[\s\S]*scroll-behavior:\s*auto\s*!important\s*;/,
    "Expected tokens.css to disable smooth scrolling (scroll-behavior: auto) under prefers-reduced-motion: reduce",
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

test("Any smooth scrolling CSS is gated behind reduced motion overrides", () => {
  const srcRoot = path.join(DESKTOP_ROOT, "src");
  const files = collectCssFiles(srcRoot);

  const offenders = [];
  const reducedMotionAttrRe = /data-reduced-motion\s*=\s*(?:"true"|'true'|true)/;
  for (const filePath of files) {
    const rawCss = fs.readFileSync(filePath, "utf8");
    const css = stripCssComments(rawCss);
    if (!/scroll-behavior\s*:\s*smooth\b/.test(css)) continue;

    // Require an explicit reduced-motion override in the same file.
    const hasReducedMotionOverride =
      /prefers-reduced-motion\s*:\s*reduce/.test(css) &&
      /scroll-behavior\s*:\s*auto\b/.test(css) &&
      reducedMotionAttrRe.test(css);

    if (!hasReducedMotionOverride) {
      offenders.push(path.relative(DESKTOP_ROOT, filePath));
    }
  }

  assert.deepEqual(
    offenders,
    [],
    `Expected any scroll-behavior: smooth usage to include reduced-motion overrides in the same file (missing in: ${offenders.join(
      ", ",
    )})`,
  );
});

test("Any smooth scrolling JS is gated behind reduced motion checks", () => {
  const srcRoot = path.join(DESKTOP_ROOT, "src");
  const files = collectSourceFiles(srcRoot);

  /** @type {string[]} */
  const offenders = [];

  const smoothBehaviorRe = /behavior\s*:\s*["']smooth["']/;
  const allowedTernaryRe = /\?\s*["']auto["']\s*:\s*["']smooth["']/;
  const reducedMotionHintRe = /prefers-reduced-motion|data-reduced-motion|getSystemReducedMotion|MEDIA\.reducedMotion/;

  for (const filePath of files) {
    const raw = fs.readFileSync(filePath, "utf8");
    const src = stripJsComments(raw);
    if (!smoothBehaviorRe.test(src)) continue;
    if (allowedTernaryRe.test(src)) continue;
    if (reducedMotionHintRe.test(src)) continue;
    offenders.push(path.relative(DESKTOP_ROOT, filePath));
  }

  assert.deepEqual(
    offenders,
    [],
    `Found scroll behavior: \"smooth\" without reduced-motion gating in: ${offenders.join(", ")}`,
  );
});

test("Sheet tabs disable smooth scrolling under reduced motion", () => {
  const uiCss = readDesktopFile("src", "styles", "ui.css");

  assert.match(
    uiCss,
    /(?:html|:root)\[data-reduced-motion\s*=\s*(?:"true"|'true'|true)\]\s+#sheet-tabs\.sheet-bar\s+\.sheet-tabs\s*\{[\s\S]*scroll-behavior:\s*auto;/,
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
