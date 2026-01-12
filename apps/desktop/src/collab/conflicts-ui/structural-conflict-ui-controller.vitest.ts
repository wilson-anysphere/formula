// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

import { StructuralConflictUiController } from "./structural-conflict-ui-controller.js";

describe("StructuralConflictUiController", () => {
  it("resolves a cell conflict via 'Keep ours' and removes it from the UI", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "c1",
      type: "cell",
      reason: "delete-vs-edit",
      sheetId: "Sheet1",
      cell: "A1",
      cellKey: "Sheet1:0:0",
      local: { kind: "edit", cellKey: "Sheet1:0:0", before: null, after: { value: 1 } },
      remote: { kind: "delete", cellKey: "Sheet1:0:0", before: { value: 2 }, after: null },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    const open = container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]');
    expect(open).not.toBeNull();
    open!.click();

    const keepOurs = container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-choose-ours"]');
    expect(keepOurs).not.toBeNull();
    keepOurs!.click();

    expect(resolveConflict).toHaveBeenCalledTimes(1);
    expect(resolveConflict).toHaveBeenCalledWith("c1", { choice: "ours" });

    expect(container.querySelector('[data-testid="structural-conflict-toast"]')).toBeNull();
    expect(container.querySelector('[data-testid="structural-conflict-dialog"]')).toBeNull();

    ui.destroy();
    container.remove();
  });

  it("resolves a move conflict via 'Use theirs destination' and removes it from the UI", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "m1",
      type: "move",
      reason: "move-destination",
      sheetId: "Sheet1",
      cell: "A1",
      cellKey: "Sheet1:0:0",
      local: { kind: "move", fromCellKey: "Sheet1:0:0", toCellKey: "Sheet1:1:0", cell: { value: 1 } },
      remote: { kind: "move", fromCellKey: "Sheet1:0:0", toCellKey: "Sheet1:2:0", cell: { value: 1 } },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    const open = container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]');
    expect(open).not.toBeNull();
    open!.click();

    const useTheirs = container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-choose-theirs"]');
    expect(useTheirs).not.toBeNull();
    useTheirs!.click();

    expect(resolveConflict).toHaveBeenCalledTimes(1);
    expect(resolveConflict).toHaveBeenCalledWith("m1", { choice: "theirs" });

    expect(container.querySelector('[data-testid="structural-conflict-toast"]')).toBeNull();
    expect(container.querySelector('[data-testid="structural-conflict-dialog"]')).toBeNull();

    ui.destroy();
    container.remove();
  });
});

