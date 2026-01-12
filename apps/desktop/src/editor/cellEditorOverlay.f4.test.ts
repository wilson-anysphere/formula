/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { CellEditorOverlay } from "./cellEditorOverlay.js";

describe("CellEditorOverlay F4 absolute reference toggle", () => {
  it("toggles the active A1 reference and keeps the token selected", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=A1");

    // Caret between A and 1.
    overlay.element.setSelectionRange(2, 2);
    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(overlay.element.value).toBe("=$A$1");
    expect(overlay.element.selectionStart).toBe(1);
    expect(overlay.element.selectionEnd).toBe(5);

    overlay.close();
    container.remove();
  });

  it("does not toggle non-formula text", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "hello");

    overlay.element.setSelectionRange(1, 1);
    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(overlay.element.value).toBe("hello");

    overlay.close();
    container.remove();
  });
});

