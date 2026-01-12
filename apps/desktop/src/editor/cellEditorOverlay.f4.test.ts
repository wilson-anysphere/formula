/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { CellEditorOverlay } from "./cellEditorOverlay.js";

describe("CellEditorOverlay F4 absolute reference toggle", () => {
  it("toggles the active A1 reference and preserves a sane caret position", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=A1");

    // Caret between A and 1.
    overlay.element.setSelectionRange(2, 2);
    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(overlay.element.value).toBe("=$A$1");
    // Caret should still be inside the reference token.
    expect(overlay.element.selectionStart).toBe(4);
    expect(overlay.element.selectionEnd).toBe(4);

    overlay.close();
    container.remove();
  });

  it("toggles when the caret is at the end of a reference token", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=A1");

    // Caret at end of token.
    overlay.element.setSelectionRange(3, 3);
    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(overlay.element.value).toBe("=$A$1");
    expect(overlay.element.selectionStart).toBe(5);
    expect(overlay.element.selectionEnd).toBe(5);

    overlay.close();
    container.remove();
  });

  it("ignores modifier chords (Alt/Ctrl/Cmd) for F4", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=A1");

    overlay.element.setSelectionRange(2, 2);

    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", altKey: true, cancelable: true }));
    expect(overlay.element.value).toBe("=A1");

    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", ctrlKey: true, cancelable: true }));
    expect(overlay.element.value).toBe("=A1");

    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", metaKey: true, cancelable: true }));
    expect(overlay.element.value).toBe("=A1");

    overlay.close();
    container.remove();
  });

  it("cycles absolute modes correctly on repeated F4 presses", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=A1");

    // Caret between A and 1.
    overlay.element.setSelectionRange(2, 2);

    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));
    expect(overlay.element.value).toBe("=$A$1");
    expect(overlay.element.selectionStart).toBe(4);
    expect(overlay.element.selectionEnd).toBe(4);

    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));
    expect(overlay.element.value).toBe("=A$1");
    expect(overlay.element.selectionStart).toBe(3);
    expect(overlay.element.selectionEnd).toBe(3);

    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));
    expect(overlay.element.value).toBe("=$A1");
    expect(overlay.element.selectionStart).toBe(3);
    expect(overlay.element.selectionEnd).toBe(3);

    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));
    expect(overlay.element.value).toBe("=A1");
    expect(overlay.element.selectionStart).toBe(2);
    expect(overlay.element.selectionEnd).toBe(2);

    overlay.close();
    container.remove();
  });

  it("preserves a full-token selection when toggling absolute refs", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=A1");

    // Select the full reference token.
    overlay.element.setSelectionRange(1, 3);
    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(overlay.element.value).toBe("=$A$1");
    // Full-token selection should expand to cover the toggled token.
    expect(overlay.element.selectionStart).toBe(1);
    expect(overlay.element.selectionEnd).toBe(5);

    overlay.close();
    container.remove();
  });

  it("does not toggle when the caret is not within a reference token", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=SUM(");

    overlay.element.setSelectionRange(overlay.element.value.length, overlay.element.value.length);
    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(overlay.element.value).toBe("=SUM(");

    overlay.close();
    container.remove();
  });

  it("does not toggle when the selection is not contained within a reference token", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const overlay = new CellEditorOverlay(container, { onCancel: () => {}, onCommit: () => {} });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "=A1+B1");

    // Selection spans A1 and the "+" operator.
    overlay.element.setSelectionRange(1, 4);
    overlay.element.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(overlay.element.value).toBe("=A1+B1");
    expect(overlay.element.selectionStart).toBe(1);
    expect(overlay.element.selectionEnd).toBe(4);

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
