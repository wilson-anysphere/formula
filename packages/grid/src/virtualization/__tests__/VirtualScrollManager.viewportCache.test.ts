import { describe, expect, it } from "vitest";
import { VirtualScrollManager } from "../VirtualScrollManager";

describe("VirtualScrollManager viewport cache", () => {
  it("returns the same object when nothing changes", () => {
    const manager = new VirtualScrollManager({ rowCount: 100, colCount: 100 });
    manager.setViewportSize(200, 200);
    manager.setFrozen(1, 1);
    manager.setScroll(50, 25);

    const first = manager.getViewportState();
    const second = manager.getViewportState();
    expect(second).toBe(first);
  });

  it("invalidates cache when scroll changes", () => {
    const manager = new VirtualScrollManager({ rowCount: 100, colCount: 100 });
    manager.setViewportSize(200, 200);
    manager.setFrozen(0, 0);
    manager.setScroll(10, 10);

    const first = manager.getViewportState();
    manager.setScroll(20, 10);
    const second = manager.getViewportState();

    expect(second).not.toBe(first);
  });

  it("invalidates cache when axis size changes", () => {
    const manager = new VirtualScrollManager({ rowCount: 100, colCount: 100 });
    manager.setViewportSize(200, 200);
    manager.setFrozen(1, 0);
    manager.setScroll(0, 0);

    const first = manager.getViewportState();

    manager.rows.setSize(0, manager.rows.defaultSize * 2);
    const second = manager.getViewportState();

    expect(second).not.toBe(first);
    expect(second.frozenHeight).toBe(manager.rows.defaultSize * 2);
  });
});

