import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);

let hasJsdom = true;
try {
  require.resolve("jsdom");
} catch {
  hasJsdom = false;
}

test(
  "Reduced motion collapses motion tokens in computed styles (jsdom)",
  { skip: !hasJsdom },
  async () => {
    const { JSDOM } = await import("jsdom");

    const desktopRoot = path.join(__dirname, "..");
    const tokensCss = fs.readFileSync(path.join(desktopRoot, "src", "styles", "tokens.css"), "utf8");
    const uiCss = fs.readFileSync(path.join(desktopRoot, "src", "styles", "ui.css"), "utf8");

    const dom = new JSDOM(
      `<!doctype html>
       <html>
         <head>
           <style>${tokensCss}\n${uiCss}</style>
         </head>
         <body>
           <div id="sheet-tabs" class="sheet-bar">
             <div class="sheet-tabs">
               <button class="sheet-tab" type="button"><span class="sheet-tab__pill">Sheet1</span></button>
             </div>
           </div>
         </body>
       </html>`,
      { pretendToBeVisual: true },
    );

    const { window } = dom;
    const { document } = window;

    const rootStyle = window.getComputedStyle(document.documentElement);
    const initialDuration = rootStyle.getPropertyValue("--motion-duration").trim();
    const initialFast = rootStyle.getPropertyValue("--motion-duration-fast").trim();

    // Sanity-check the values are actually present (avoid passing if jsdom doesn't
    // support custom properties at all).
    assert.notEqual(initialDuration, "");
    assert.notEqual(initialFast, "");

    document.documentElement.setAttribute("data-reduced-motion", "true");

    const reducedStyle = window.getComputedStyle(document.documentElement);
    assert.equal(reducedStyle.getPropertyValue("--motion-duration").trim(), "0ms");
    assert.equal(reducedStyle.getPropertyValue("--motion-duration-fast").trim(), "0ms");
    assert.equal(reducedStyle.getPropertyValue("--motion-ease").trim(), "linear");

    // Optional: jsdom's CSS support for `scroll-behavior` can vary by version. If it
    // computes, verify reduced motion disables the smooth behavior.
    const sheetTabs = document.querySelector(".sheet-tabs");
    assert.ok(sheetTabs, "Expected .sheet-tabs element to exist");
    const scrollBehavior = window.getComputedStyle(sheetTabs).getPropertyValue("scroll-behavior").trim();
    if (scrollBehavior) {
      assert.equal(scrollBehavior, "auto");
    }
  },
);
