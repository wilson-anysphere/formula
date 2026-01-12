/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { CellEditorOverlay } from "./cellEditorOverlay.js";

describe("CellEditorOverlay programmatic commit/cancel", () => {
  it("commits the current value with a command reason and closes the editor", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const onCommit = vi.fn();
    const onCancel = vi.fn();

    const overlay = new CellEditorOverlay(container, { onCommit, onCancel });
    overlay.open({ row: 0, col: 0 }, { x: 0, y: 0, width: 100, height: 24 }, "initial");

    overlay.element.value = "next";
    overlay.commit("command");

    expect(onCancel).not.toHaveBeenCalled();
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith({
      cell: { row: 0, col: 0 },
      value: "next",
      reason: "command",
      shift: false,
    });

    expect(overlay.isOpen()).toBe(false);
    container.remove();
  });

  it("cancels the active edit and closes the editor", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const onCommit = vi.fn();
    const onCancel = vi.fn();

    const overlay = new CellEditorOverlay(container, { onCommit, onCancel });
    overlay.open({ row: 1, col: 2 }, { x: 0, y: 0, width: 100, height: 24 }, "hello");

    overlay.cancel();

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(onCancel).toHaveBeenCalledWith({ row: 1, col: 2 });
    expect(overlay.isOpen()).toBe(false);
    container.remove();
  });
});

