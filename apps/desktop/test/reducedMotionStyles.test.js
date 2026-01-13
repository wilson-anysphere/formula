import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function readDesktopFile(...segments) {
  return fs.readFileSync(path.join(__dirname, "..", ...segments), "utf8");
}

function extractTransitionDeclarations(css) {
  return [...css.matchAll(/transition\s*:\s*([\s\S]*?);/g)].map((match) => match[1]?.trim() ?? "");
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

test("Core UI surfaces use motion tokens for transitions (so reduced motion disables animation)", () => {
  const files = [
    ["src", "styles", "ribbon.css"],
    ["src", "styles", "dialogs.css"],
    ["src", "styles", "context-menu.css"],
    ["src", "styles", "ui.css"],
  ];

  for (const relPath of files) {
    const css = readDesktopFile(...relPath);

    assert.equal(
      /\banimation\s*:/.test(css) || /@keyframes\b/.test(css),
      false,
      `Unexpected CSS animations in ${relPath.join("/")}; animations must be gated behind reduced-motion`,
    );

    for (const transition of extractTransitionDeclarations(css)) {
      assert.match(
        transition,
        /var\(--motion-duration(?:-fast)?\)/,
        `Expected transition declarations in ${relPath.join("/")} to use --motion-duration tokens`,
      );
      assert.match(
        transition,
        /var\(--motion-ease\)/,
        `Expected transition declarations in ${relPath.join("/")} to use --motion-ease token`,
      );
    }

    // Guard against accidentally hardcoding durations in leaf CSS files (token definitions live in tokens.css).
    if (!relPath.includes("tokens.css")) {
      assert.equal(
        /\b\d+(?:\.\d+)?(?:ms|s)\b/.test(css),
        false,
        `Expected ${relPath.join("/")} to avoid hard-coded time values; use motion tokens instead`,
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
