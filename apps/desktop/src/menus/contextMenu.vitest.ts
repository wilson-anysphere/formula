// @vitest-environment jsdom

import { describe, it, expect, vi } from "vitest";

import { ContextMenu } from "./contextMenu.js";

describe("ContextMenu (DOM)", () => {
  it("renders items as native <button> elements without role overrides", () => {
    const menu = new ContextMenu();
    try {
      menu.open({
        x: 0,
        y: 0,
        items: [
          {
            type: "item",
            label: "Test Item",
            onSelect: () => {},
          },
        ],
      });

      const overlay = document.querySelector('[data-testid="context-menu"]');
      expect(overlay, "expected context menu overlay to be attached").toBeTruthy();
      expect(overlay?.getAttribute("data-keybinding-barrier")).toBe("true");

      const container = overlay?.querySelector('[role="menu"]');
      expect(container, "expected menu container to have role=menu").toBeTruthy();

      const button = overlay?.querySelector("button");
      expect(button, "expected menu item to render as a <button>").toBeInstanceOf(HTMLButtonElement);
      expect(button?.getAttribute("role"), "menu items should keep the implicit button role").toBeNull();

      expect(
        overlay?.querySelector('[role="menuitem"]'),
        "context menu should not override item buttons to role=menuitem (Playwright selectors depend on role=button)",
      ).toBeNull();
    } finally {
      menu.close();
    }
  });

  it("ignores outside scroll events immediately after open (prevents instant-dismiss flake)", () => {
    const menu = new ContextMenu();
    const scroller = document.createElement("div");
    document.body.appendChild(scroller);

    // Deterministic timing: ContextMenu uses `performance.now()` to set a short grace period.
    let now = 0;
    const nowSpy = vi.spyOn(performance, "now").mockImplementation(() => now);

    try {
      menu.open({
        x: 0,
        y: 0,
        items: [
          {
            type: "item",
            label: "Test Item",
            onSelect: () => {},
          },
        ],
      });

      expect(menu.isOpen()).toBe(true);

      // Simulate a scroll event occurring right after the menu opens (e.g. focus-induced scrollIntoView).
      scroller.dispatchEvent(new Event("scroll", { bubbles: true }));
      expect(menu.isOpen()).toBe(true);

      // After the grace period, outside scroll events should close the menu.
      now = 200;
      scroller.dispatchEvent(new Event("scroll", { bubbles: true }));
      expect(menu.isOpen()).toBe(false);
    } finally {
      nowSpy.mockRestore();
      scroller.remove();
      menu.close();
    }
  });
});
