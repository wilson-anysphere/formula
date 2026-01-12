// @vitest-environment jsdom

import { describe, it, expect } from "vitest";

import { ContextMenu } from "./contextMenu.js";

describe("ContextMenu (DOM)", () => {
  it("renders items as native <button> elements without role overrides", () => {
    const menu = new ContextMenu();
    try {
      menu.open({
        x: 0,
        y: 0,
        focusFirst: true,
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
});

