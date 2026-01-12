// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

import { StructuralConflictUiController } from "./structural-conflict-ui-controller.js";

describe("StructuralConflictUiController", () => {
  it("renders conflict locations using sheet display names when sheetNameResolver is provided", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const namesById = new Map<string, string>([["sheet_123", "My Sheet"]]);
    const sheetNameResolver = {
      getSheetNameById: (id: string) => namesById.get(id) ?? null,
      getSheetIdByName: (_name: string) => null,
    };

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
      sheetNameResolver,
    });

    ui.addConflict({
      id: "c_display",
      type: "cell",
      reason: "content",
      sheetId: "sheet_123",
      cell: "A1",
      cellKey: "sheet_123:0:0",
      local: { kind: "edit", cellKey: "sheet_123:0:0", before: null, after: { value: 1 } },
      remote: { kind: "edit", cellKey: "sheet_123:0:0", before: null, after: { value: 2 } },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    const toast = container.querySelector<HTMLElement>('[data-testid="structural-conflict-toast"]');
    expect(toast).not.toBeNull();
    expect(toast!.textContent).toContain("'My Sheet'!A1");
    expect(toast!.textContent).not.toContain("sheet_123");

    ui.destroy();
    container.remove();
  });

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

  it("supports manual resolution for cell conflicts", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "c2",
      type: "cell",
      reason: "content",
      sheetId: "Sheet1",
      cell: "B2",
      cellKey: "Sheet1:1:1",
      local: { kind: "edit", cellKey: "Sheet1:1:1", before: null, after: { value: 1 } },
      remote: { kind: "edit", cellKey: "Sheet1:1:1", before: null, after: { value: 2 } },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]')!.click();
    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-manual"]')!.click();

    const textarea = container.querySelector<HTMLTextAreaElement>('[data-testid="structural-conflict-manual-cell"]');
    expect(textarea).not.toBeNull();
    textarea!.value = JSON.stringify({ value: "manual" });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-manual-apply"]')!.click();

    expect(resolveConflict).toHaveBeenCalledWith("c2", { choice: "manual", cell: { value: "manual" } });
    expect(container.querySelector('[data-testid="structural-conflict-toast"]')).toBeNull();

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

  it("supports manual resolution for move conflicts", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "m2",
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

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]')!.click();
    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-manual"]')!.click();

    const dest = container.querySelector<HTMLInputElement>('[data-testid="structural-conflict-manual-destination"]');
    expect(dest).not.toBeNull();
    dest!.value = "Sheet1:5:5";

    const textarea = container.querySelector<HTMLTextAreaElement>('[data-testid="structural-conflict-manual-cell"]');
    expect(textarea).not.toBeNull();
    textarea!.value = JSON.stringify({ value: 42 });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-manual-apply"]')!.click();

    expect(resolveConflict).toHaveBeenCalledWith("m2", { choice: "manual", to: "Sheet1:5:5", cell: { value: 42 } });
    expect(container.querySelector('[data-testid="structural-conflict-toast"]')).toBeNull();

    ui.destroy();
    container.remove();
  });
});
