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
      listDrawingsForSheet: vi.fn(() => drawings),
      isSelectedDrawingImage: vi.fn(() => true),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      duplicateSelectedDrawing: vi.fn(),
      deleteDrawingById: vi.fn((id: number) => {
        drawings = drawings.filter((d) => d.id !== id);
        if (selectedId === id) selectedId = null;
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
      expect(labels).toEqual(["Cut", "Copy", "Duplicate", "Delete", "Bring Forward", "Send Backward"]);

      // Ensure we didn't leak cell-oriented items into the drawing menu.
      expect(labels).not.toContain("Paste");

      const deleteBtn = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []).find((btn) => {
        const text = btn.querySelector<HTMLElement>(".context-menu__label")?.textContent ?? "";
        return text.trim() === "Delete";
      });
      expect(deleteBtn).toBeTruthy();
      deleteBtn!.click();

      expect(app.deleteDrawingById).toHaveBeenCalledTimes(1);
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
      listDrawingsForSheet: vi.fn(() => [{ id: 1 }, { id: 2 }]),
      isSelectedDrawingImage: vi.fn(() => true),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      duplicateSelectedDrawing: vi.fn(),
      deleteDrawingById: vi.fn(),
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

  it("disables Cut/Copy for non-image drawings", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-non-image" });
    let selectedId: number | null = 1;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      listDrawingsForSheet: vi.fn(() => [{ id: 2 }, { id: 1 }, { id: 3 }]),
      isSelectedDrawingImage: vi.fn(() => false),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      duplicateSelectedDrawing: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-non-image"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      expect(buttonByLabel("Cut")?.disabled).toBe(true);
      expect(buttonByLabel("Copy")?.disabled).toBe(true);
      expect(buttonByLabel("Duplicate")?.disabled).toBe(false);
      expect(buttonByLabel("Delete")?.disabled).toBe(false);
      expect(buttonByLabel("Bring Forward")?.disabled).toBe(false);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(false);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-non-image"]')?.remove();
    }
  });

  it("disables z-order actions when the selected drawing is already topmost/backmost", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-z-order" });
    let selectedId: number | null = 1;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      // Topmost-first ordering: selected id=1 is topmost.
      listDrawingsForSheet: vi.fn(() => [{ id: 1 }, { id: 2 }]),
      isSelectedDrawingImage: vi.fn(() => true),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-z-order"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      expect(buttonByLabel("Bring Forward")?.disabled).toBe(true);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(false);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-z-order"]')?.remove();
    }
  });

  it("disables delete/cut/arrange in read-only mode but allows Copy for image drawings", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-read-only" });
    let selectedId: number | null = 1;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 1 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      listDrawingsForSheet: vi.fn(() => [{ id: 2 }, { id: 1 }]),
      isSelectedDrawingImage: vi.fn(() => true),
      isReadOnly: vi.fn(() => true),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-read-only"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      expect(buttonByLabel("Cut")?.disabled).toBe(true);
      expect(buttonByLabel("Copy")?.disabled).toBe(false);
      expect(buttonByLabel("Delete")?.disabled).toBe(true);
      expect(buttonByLabel("Bring Forward")?.disabled).toBe(true);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(true);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-read-only"]')?.remove();
    }
  });

  it("disables z-order actions for a single canvas chart (no reorder possible)", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-chart" });
    let selectedId: number | null = -123;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: -123 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      listDrawingsForSheet: vi.fn(() => [{ id: -123 }, { id: 1 }]),
      isSelectedDrawingImage: vi.fn(() => false),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-chart"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      expect(buttonByLabel("Bring Forward")?.disabled).toBe(true);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(true);
      expect(buttonByLabel("Delete")?.disabled).toBe(false);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-chart"]')?.remove();
    }
  });

  it("enables z-order actions for canvas charts within their chart stack", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-chart-stack" });
    let selectedId: number | null = -2;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: -2 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      // Topmost-first ordering: charts first, then workbook drawings.
      listDrawingsForSheet: vi.fn(() => [{ id: -1 }, { id: -2 }, { id: -3 }, { id: 1 }]),
      isSelectedDrawingImage: vi.fn(() => false),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-chart-stack"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      expect(buttonByLabel("Bring Forward")?.disabled).toBe(false);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(false);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-chart-stack"]')?.remove();
    }
  });

  it("treats large-magnitude negative ids as workbook drawings (not canvas charts) for z-order gating", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-hashed-id" });
    const hashedDrawingId = -0x200000000; // 2^33 (see parseDrawingObjectId in drawings/modelAdapters.ts)
    let selectedId: number | null = hashedDrawingId;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: hashedDrawingId })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      // Topmost-first ordering: ChartStore chart ids are negative 32-bit; hashed drawing ids are
      // a separate large-magnitude negative namespace and should behave like drawings.
      listDrawingsForSheet: vi.fn(() => [{ id: -1 }, { id: hashedDrawingId }, { id: 2 }]),
      isSelectedDrawingImage: vi.fn(() => false),
      isReadOnly: vi.fn(() => false),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-hashed-id"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      // `hashedDrawingId` is the topmost workbook drawing but cannot be moved above charts.
      expect(buttonByLabel("Bring Forward")?.disabled).toBe(true);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(false);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-hashed-id"]')?.remove();
    }
  });

  it("disables Bring Forward for the topmost workbook drawing when canvas charts exist", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-z-order-with-charts" });
    let selectedId: number | null = 10;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: 10 })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      // Topmost-first ordering: charts first, then workbook drawings. Drawing id=10 is topmost
      // within the drawings stack, even though it's not index 0 overall.
      listDrawingsForSheet: vi.fn(() => [{ id: -1 }, { id: -2 }, { id: 10 }, { id: 11 }]),
      isSelectedDrawingImage: vi.fn(() => false),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-z-order-with-charts"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      expect(buttonByLabel("Bring Forward")?.disabled).toBe(true);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(false);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-z-order-with-charts"]')?.remove();
    }
  });

  it("enables z-order actions for hashed (far-negative) drawing ids within the workbook drawings stack", () => {
    const contextMenu = new ContextMenu({ testId: "context-menu-drawing-hashed-id-mid-stack" });
    const hashedId = -0x200000000; // -2^33 (hashed drawing id namespace floor)
    let selectedId: number | null = hashedId;
    const app = {
      hitTestDrawingAtClientPoint: vi.fn(() => ({ id: hashedId })),
      getSelectedDrawingId: vi.fn(() => selectedId),
      // Topmost-first ordering: ChartStore charts first (negative 32-bit), then workbook drawings.
      // The hashed drawing id is mid-stack within the workbook drawings group, so it should be able
      // to move in either direction *within that group*.
      listDrawingsForSheet: vi.fn(() => [{ id: -1 }, { id: -2 }, { id: 1 }, { id: hashedId }, { id: 2 }]),
      isSelectedDrawingImage: vi.fn(() => false),
      isReadOnly: vi.fn(() => false),
      selectDrawingById: vi.fn((id: number | null) => {
        selectedId = id;
      }),
      cut: vi.fn(),
      copy: vi.fn(),
      deleteDrawingById: vi.fn(),
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
        isEditing: false,
      });

      const overlay = document.querySelector<HTMLElement>('[data-testid="context-menu-drawing-hashed-id-mid-stack"]');
      const buttons = Array.from(overlay?.querySelectorAll<HTMLButtonElement>("button") ?? []);
      const buttonByLabel = (label: string) =>
        buttons.find((btn) => (btn.querySelector(".context-menu__label")?.textContent ?? "").trim() === label) ?? null;

      expect(buttonByLabel("Bring Forward")?.disabled).toBe(false);
      expect(buttonByLabel("Send Backward")?.disabled).toBe(false);
    } finally {
      contextMenu.close();
      document.querySelector('[data-testid="context-menu-drawing-hashed-id-mid-stack"]')?.remove();
    }
  });
});
