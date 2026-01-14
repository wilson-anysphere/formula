/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { CellEditorOverlay } from "./cellEditorOverlay.js";

describe("CellEditorOverlay IME composition safety", () => {
  it("does not commit on Enter during composition, but does after compositionend", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const overlay = new CellEditorOverlay(container, { onCommit, onCancel });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "hello");

    overlay.element.dispatchEvent(new Event("compositionstart"));
    const enterDuringComposition = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    overlay.element.dispatchEvent(enterDuringComposition);

    expect(enterDuringComposition.defaultPrevented).toBe(false);
    expect(onCommit).not.toHaveBeenCalled();
    expect(overlay.isOpen()).toBe(true);

    overlay.element.dispatchEvent(new Event("compositionend"));
    const enterAfterComposition = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    overlay.element.dispatchEvent(enterAfterComposition);

    expect(enterAfterComposition.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit.mock.calls[0]?.[0]).toEqual({
      cell: { row: 0, col: 0 },
      value: "hello",
      reason: "enter",
      shift: false,
    });
    expect(onCancel).not.toHaveBeenCalled();
    expect(overlay.isOpen()).toBe(false);

    container.remove();
  });

  it("does not commit on Tab during composition (prevents focus traversal), but does after compositionend", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const overlay = new CellEditorOverlay(container, { onCommit, onCancel });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "hello");

    overlay.element.dispatchEvent(new Event("compositionstart"));
    const tabDuringComposition = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    overlay.element.dispatchEvent(tabDuringComposition);

    expect(tabDuringComposition.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).not.toHaveBeenCalled();
    expect(overlay.isOpen()).toBe(true);

    overlay.element.dispatchEvent(new Event("compositionend"));
    const tabAfterComposition = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    overlay.element.dispatchEvent(tabAfterComposition);

    expect(tabAfterComposition.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit.mock.calls[0]?.[0]).toEqual({
      cell: { row: 0, col: 0 },
      value: "hello",
      reason: "tab",
      shift: false,
    });
    expect(onCancel).not.toHaveBeenCalled();
    expect(overlay.isOpen()).toBe(false);

    container.remove();
  });
});
