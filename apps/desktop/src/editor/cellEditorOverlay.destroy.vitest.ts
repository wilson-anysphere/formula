/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { CellEditorOverlay } from "./cellEditorOverlay.js";

describe("CellEditorOverlay destroy", () => {
  it("removes the element and prevents future edits after destroy()", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCommit: vi.fn(), onCancel: vi.fn() });
    expect(overlay.element.isConnected).toBe(true);

    overlay.destroy();

    expect(overlay.element.isConnected).toBe(false);
    expect(container.contains(overlay.element)).toBe(false);
    expect(overlay.isOpen()).toBe(false);

    // No-op after destroy.
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "hello");
    expect(overlay.isOpen()).toBe(false);
    expect(overlay.element.isConnected).toBe(false);

    container.remove();
  });
});

