import { describe, expect, it } from "vitest";
import { VirtualScrollManager } from "../VirtualScrollManager";

describe("VirtualScrollManager", () => {
  it("calculates main visible ranges with frozen panes", () => {
    const manager = new VirtualScrollManager({
      rowCount: 1000,
      colCount: 1000,
      defaultRowHeight: 10,
      defaultColWidth: 50
    });

    manager.setViewportSize(120, 45);
    manager.setFrozen(2, 1);
    manager.setScroll(0, 0);

    const viewport = manager.getViewportState();
    expect(viewport.frozenWidth).toBe(50);
    expect(viewport.frozenHeight).toBe(20);

    expect(viewport.main.rows.start).toBe(2);
    expect(viewport.main.rows.offset).toBe(0);
    expect(viewport.main.rows.end).toBe(5);

    expect(viewport.main.cols.start).toBe(1);
    expect(viewport.main.cols.offset).toBe(0);
    expect(viewport.main.cols.end).toBe(3);
  });

  it("accounts for partial scroll offsets", () => {
    const manager = new VirtualScrollManager({
      rowCount: 1000,
      colCount: 1000,
      defaultRowHeight: 10,
      defaultColWidth: 50
    });

    manager.setViewportSize(120, 45);
    manager.setFrozen(2, 1);
    manager.setScroll(0, 5);

    const viewport = manager.getViewportState();
    expect(viewport.main.rows.start).toBe(2);
    expect(viewport.main.rows.offset).toBe(5);
    expect(viewport.main.rows.end).toBe(5);
  });
});

