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
    const ribbonCss = fs.readFileSync(path.join(desktopRoot, "src", "styles", "ribbon.css"), "utf8");
    const dialogsCss = fs.readFileSync(path.join(desktopRoot, "src", "styles", "dialogs.css"), "utf8");
    const contextMenuCss = fs.readFileSync(path.join(desktopRoot, "src", "styles", "context-menu.css"), "utf8");

    const dom = new JSDOM(
      `<!doctype html>
       <html>
         <head>
           <style>${tokensCss}\n${uiCss}\n${ribbonCss}\n${dialogsCss}\n${contextMenuCss}</style>
         </head>
         <body>
           <div class="ribbon">
             <button type="button" class="ribbon__tab">Home</button>
           </div>
 
           <dialog class="dialog" open>
             <div class="dialog__controls">
               <button type="button">OK</button>
             </div>
           </dialog>
 
           <div class="context-menu-overlay">
             <div class="context-menu">
               <button type="button" class="context-menu__item"><span class="context-menu__label">Copy</span></button>
             </div>
           </div>
 
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

    const assertTransitionCollapsed = (selector, label) => {
      const el = document.querySelector(selector);
      assert.ok(el, `Expected ${label} element (${selector}) to exist`);
      const duration = window.getComputedStyle(el).getPropertyValue("transition-duration").trim();
      if (!duration) return; // jsdom might not compute transition shorthands
      if (duration.includes("var(")) return; // older jsdom may not resolve custom properties here
      const parts = duration.split(",").map((p) => p.trim());
      for (const part of parts) {
        assert.ok(
          part === "0s" || part === "0ms",
          `Expected ${label} transition-duration to collapse to 0ms under reduced motion (got: ${duration})`,
        );
      }
    };
 
    // Verify that key UI surfaces (ribbon, dialogs, context menu) do not animate once
    // reduced motion is enabled (when jsdom provides transition-duration values).
    assertTransitionCollapsed(".ribbon__tab", "ribbon tab");
    assertTransitionCollapsed(".dialog__controls button", "dialog button");
    assertTransitionCollapsed(".context-menu__item", "context menu item");
 
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
