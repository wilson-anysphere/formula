// @vitest-environment jsdom

import { describe, it, expect, vi } from "vitest";

import { ContextMenu } from "./menus/contextMenu.js";
import { tryOpenDrawingContextMenuAtClientPoint } from "./mainContextMenuDrawing.js";

describe("Drawing context menu (main wiring helper)", () => {
  it("opens a picture context menu on right-click and deletes the drawing", () => {
    const gridRoot = document.createElement("div");
    document.body.appendChild(gridRoot);

    const contextMenu = new ContextMenu({ testId: "context-menu-drawing" });

    let drawings = [{ id: 1 }];
    let selectedId: number | null = null;

    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteSelectedDrawing: vi.fn(() => {
        if (selectedId == null) return;
        drawings = drawings.filter((d) => d.id !== selectedId);
        selectedId = null;
      }),
      bringSelectedDrawingForward: vi.fn(),
      sendSelectedDrawingBackward: vi.fn(),
      focus: vi.fn(),
    } as any;

    gridRoot.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      tryOpenDrawingContextMenuAtClientPoint({
        app,
        contextMenu,
        clientX: e.clientX,
        clientY: e.clientY,
        isEditing: false,
      });
    });

    try {
      gridRoot.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, clientX: 10, clientY: 20 }));

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing"]');
      expect(overlay).toBeTruthy();
      expect(overlay?.hidden).toBe(false);

      const labels = Array.from(overlay?.querySelectorAll<HTMLElement>(".context-menu__label") ?? []).map((el) =>
        (el.textContent ?? "").trim(),
      );
      expect(labels).toEqual(["Cut", "Copy", "Delete", "Bring Forward", "Send Backward"]);

      // Ensure we didn't leak cell-oriented items into the drawing menu.
      expect(labels).not.toContain("Paste");

      const deleteBtn = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []).find((btn) => {
        const text = btn.querySelector<HTMLElement>(".context-menu__label")?.textContent ?? "";
        return text.trim() === "Delete";
      });
      expect(deleteBtn).toBeTruthy();
      deleteBtn!.click();

      expect(app.deleteSelectedDrawing).toHaveBeenCalledTimes(1);
      expect(drawings).toEqual([]);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing"]')?.remove();
      gridRoot.remove();
    }
  });

  it("disables drawing actions while editing", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-editing" });
    let selectedId: number | null = 1;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteSelectedDrawing: vi.fn(),
      bringSelectedDrawingForward: vi.fn(),
      sendSelectedDrawingBackward: vi.fn(),
      focus: vi.fn(),
    } as any;

    try {
      tryOpenDrawingContextMenuAtClientPoint({
        app,
        contextMenu,
        clientX: 10,
        clientY: 20,
        isEditing: true,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-editing"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      expect(buttons.length).toBeGreaterThan(0);
      for (const btn of buttons) {
        expect(btn.disabled).toBe(true);
      }
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-editing"]')?.remove();
    }
  });
});
