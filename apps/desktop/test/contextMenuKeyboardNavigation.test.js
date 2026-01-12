import test from "node:test";
import assert from "node:assert/strict";

import { ContextMenu } from "../src/menus/contextMenu.ts";

let JSDOM = null;
try {
  // `jsdom` is optional for lightweight node:test runs (some agent environments do not
  // install workspace dev dependencies). Skip DOM-specific tests when unavailable.
  // eslint-disable-next-line node/no-unsupported-features/es-syntax
  ({ JSDOM } = await import("jsdom"));
} catch {
  // ignore
}

const hasDom = Boolean(JSDOM);

function withDom(fn) {
  const dom = new JSDOM("<!doctype html><html><body></body></html>", { url: "http://localhost" });

  /** @type {Record<string, any>} */
  const prev = {
    window: globalThis.window,
    document: globalThis.document,
    Node: globalThis.Node,
    HTMLElement: globalThis.HTMLElement,
    HTMLDivElement: globalThis.HTMLDivElement,
    HTMLButtonElement: globalThis.HTMLButtonElement,
  };

  globalThis.window = dom.window;
  globalThis.document = dom.window.document;
  globalThis.Node = dom.window.Node;
  globalThis.HTMLElement = dom.window.HTMLElement;
  globalThis.HTMLDivElement = dom.window.HTMLDivElement;
  globalThis.HTMLButtonElement = dom.window.HTMLButtonElement;

  try {
    fn(dom);
  } finally {
    for (const [key, value] of Object.entries(prev)) {
      if (value === undefined) {
        delete globalThis[key];
      } else {
        globalThis[key] = value;
      }
    }
  }
}

test(
  "ContextMenu.update preserves focus when called while a submenu is focused",
  { skip: !hasDom },
  () => {
  withDom((dom) => {
    const menu = new ContextMenu();

    const items = [
      {
        type: "item",
        label: "Cut",
        onSelect: () => {
          // ignore
        },
      },
      {
        type: "submenu",
        label: "Format",
        items: [
          {
            type: "item",
            label: "Bold",
            onSelect: () => {
              // ignore
            },
          },
        ],
      },
    ];

    try {
      menu.open({ x: 10, y: 10, items });
      assert.equal(dom.window.document.activeElement?.textContent, "Cut");

      dom.window.dispatchEvent(new dom.window.KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true }));
      assert.ok(String(dom.window.document.activeElement?.textContent).includes("Format"));

      dom.window.dispatchEvent(new dom.window.KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
      assert.ok(dom.window.document.querySelector(".context-menu__submenu"));
      assert.ok(String(dom.window.document.activeElement?.textContent).includes("Bold"));

      // Simulate ContextKeyService-driven updates while focus is inside a submenu.
      menu.update(items);

      assert.equal(dom.window.document.querySelectorAll(".context-menu__submenu").length, 0);
      assert.ok(dom.window.document.activeElement instanceof dom.window.HTMLButtonElement);
      assert.equal(dom.window.document.activeElement.querySelector(".context-menu__label")?.textContent, "Format");
    } finally {
      menu.close();
    }
  });
  },
);

test("ContextMenu closes submenu and restores focus when the main menu scrolls", { skip: !hasDom }, () => {
  withDom((dom) => {
    const menu = new ContextMenu();

    const items = [
      {
        type: "submenu",
        label: "Format",
        items: [
          {
            type: "item",
            label: "Bold",
            onSelect: () => {
              // ignore
            },
          },
        ],
      },
    ];

    try {
      menu.open({ x: 10, y: 10, items });

      dom.window.dispatchEvent(new dom.window.KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
      assert.ok(dom.window.document.querySelector(".context-menu__submenu"));
      assert.ok(String(dom.window.document.activeElement?.textContent).includes("Bold"));

      const mainMenu = dom.window.document.querySelector(".context-menu");
      assert.ok(mainMenu);

      mainMenu.dispatchEvent(new dom.window.Event("scroll"));

      assert.equal(dom.window.document.querySelectorAll(".context-menu__submenu").length, 0);
      assert.ok(dom.window.document.activeElement instanceof dom.window.HTMLButtonElement);
      assert.equal(dom.window.document.activeElement.querySelector(".context-menu__label")?.textContent, "Format");
    } finally {
      menu.close();
    }
  });
});
