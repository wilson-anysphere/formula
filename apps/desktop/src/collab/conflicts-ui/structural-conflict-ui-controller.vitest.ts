// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import * as Y from "yjs";

import { StructuralConflictUiController } from "./structural-conflict-ui-controller.js";
import { CellStructuralConflictMonitor } from "../../../../../packages/collab/conflicts/index.js";

describe("StructuralConflictUiController", () => {
  it("renders a Jump to cell button and invokes the callback with the conflict cell", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const onNavigateToCell = vi.fn();
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
      onNavigateToCell,
    });

    ui.addConflict({
      id: "c_jump",
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

    const jump = container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-jump-to-cell"]');
    expect(jump).not.toBeNull();
    jump!.click();

    expect(onNavigateToCell).toHaveBeenCalledTimes(1);
    expect(onNavigateToCell).toHaveBeenCalledWith({ sheetId: "Sheet1", row: 1, col: 1 });

    ui.destroy();
    container.remove();
  });

  it("ignores errors thrown by onNavigateToCell", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const onNavigateToCell = vi.fn(() => {
      throw new Error("boom");
    });

    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
      onNavigateToCell,
    });

    ui.addConflict({
      id: "c_jump_throw",
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

    const jump = container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-jump-to-cell"]');
    expect(jump).not.toBeNull();
    expect(() => jump!.click()).not.toThrow();
    expect(onNavigateToCell).toHaveBeenCalledWith({ sheetId: "Sheet1", row: 1, col: 1 });

    ui.destroy();
    container.remove();
  });

  it("applies a user label resolver when rendering the remote (theirs) panel label", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
      resolveUserLabel: (userId: string) => (userId === "u2" ? "  Bob  " : userId),
    });

    ui.addConflict({
      id: "c_label",
      type: "cell",
      reason: "content",
      sheetId: "Sheet1",
      cell: "A1",
      cellKey: "Sheet1:0:0",
      local: { kind: "edit", cellKey: "Sheet1:0:0", before: null, after: { value: 1 } },
      remote: { kind: "edit", cellKey: "Sheet1:0:0", before: null, after: { value: 2 } },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]')!.click();

    const remotePanel = container.querySelector<HTMLElement>('[data-testid="structural-conflict-remote"]');
    expect(remotePanel).not.toBeNull();
    const label = remotePanel!.querySelector<HTMLElement>(".conflict-dialog__panel-label");
    expect(label?.textContent).toBe("Theirs (Bob)");

    ui.destroy();
    container.remove();
  });

  it("falls back to remote user id when resolveUserLabel throws", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
      resolveUserLabel: () => {
        throw new Error("boom");
      },
    });

    ui.addConflict({
      id: "c_label_throw",
      type: "cell",
      reason: "content",
      sheetId: "Sheet1",
      cell: "A1",
      cellKey: "Sheet1:0:0",
      local: { kind: "edit", cellKey: "Sheet1:0:0", before: null, after: { value: 1 } },
      remote: { kind: "edit", cellKey: "Sheet1:0:0", before: null, after: { value: 2 } },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]')!.click();

    const remotePanel = container.querySelector<HTMLElement>('[data-testid="structural-conflict-remote"]');
    expect(remotePanel).not.toBeNull();
    const label = remotePanel!.querySelector<HTMLElement>(".conflict-dialog__panel-label");
    expect(label?.textContent).toBe("Theirs (u2)");

    ui.destroy();
    container.remove();
  });

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

  it("renders a token diff view when structural conflict ops include formulas", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const ui = new StructuralConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "c_formula",
      type: "cell",
      reason: "content",
      sheetId: "Sheet1",
      cell: "A1",
      cellKey: "Sheet1:0:0",
      local: { kind: "edit", cellKey: "Sheet1:0:0", before: null, after: { formula: "=A1" } },
      remote: { kind: "edit", cellKey: "Sheet1:0:0", before: null, after: { formula: "=A2" } },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]')!.click();

    const diff = container.querySelector<HTMLElement>('[data-testid="structural-conflict-formula-diff"]');
    expect(diff).not.toBeNull();
    expect(diff!.querySelector(".formula-diff-op--delete")).not.toBeNull();
    expect(diff!.querySelector(".formula-diff-op--insert")).not.toBeNull();

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

  it("integrates with CellStructuralConflictMonitor (cell conflict)", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const needsCryptoStub = typeof (globalThis as any).crypto?.randomUUID !== "function";
    if (needsCryptoStub) {
      vi.stubGlobal("crypto", { randomUUID: () => `uuid-${Math.random().toString(16).slice(2)}` } as any);
    }

    const doc = new Y.Doc();
    let ui: StructuralConflictUiController | null = null;
    const monitor = new CellStructuralConflictMonitor({
      doc,
      localUserId: "u1",
      onConflict: (conflict: any) => ui?.addConflict(conflict),
    });

    ui = new StructuralConflictUiController({ container, monitor });

    (monitor as any)._emitConflict({
      type: "cell",
      reason: "content",
      sourceCellKey: "Sheet1:0:0",
      local: { kind: "edit", userId: "u1", cellKey: "Sheet1:0:0", before: null, after: { value: 123 } },
      remote: { kind: "edit", userId: "u2", cellKey: "Sheet1:0:0", before: null, after: { value: 456 } },
    });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]')!.click();
    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-choose-ours"]')!.click();

    const ycell = doc.getMap("cells").get("Sheet1:0:0") as any;
    expect(ycell?.get?.("value")).toBe(123);
    expect(monitor.listConflicts().length).toBe(0);

    ui.destroy();
    monitor.dispose();
    container.remove();
    if (needsCryptoStub) vi.unstubAllGlobals();
  });

  it("integrates with CellStructuralConflictMonitor (move conflict)", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const needsCryptoStub = typeof (globalThis as any).crypto?.randomUUID !== "function";
    if (needsCryptoStub) {
      vi.stubGlobal("crypto", { randomUUID: () => `uuid-${Math.random().toString(16).slice(2)}` } as any);
    }

    const doc = new Y.Doc();
    let ui: StructuralConflictUiController | null = null;
    const monitor = new CellStructuralConflictMonitor({
      doc,
      localUserId: "u1",
      onConflict: (conflict: any) => ui?.addConflict(conflict),
    });

    ui = new StructuralConflictUiController({ container, monitor });

    (monitor as any)._emitConflict({
      type: "move",
      reason: "move-destination",
      sourceCellKey: "Sheet1:0:0",
      local: { kind: "move", userId: "u1", fromCellKey: "Sheet1:0:0", toCellKey: "Sheet1:0:1", cell: { value: 1 } },
      remote: { kind: "move", userId: "u2", fromCellKey: "Sheet1:0:0", toCellKey: "Sheet1:0:2", cell: { value: 2 } },
    });

    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-toast-open"]')!.click();
    container.querySelector<HTMLButtonElement>('[data-testid="structural-conflict-choose-ours"]')!.click();

    const cells = doc.getMap("cells") as any;
    expect(cells.get("Sheet1:0:0")).toBeUndefined();
    expect(cells.get("Sheet1:0:2")).toBeUndefined();
    const dest = cells.get("Sheet1:0:1") as any;
    expect(dest?.get?.("value")).toBe(1);
    expect(monitor.listConflicts().length).toBe(0);

    ui.destroy();
    monitor.dispose();
    container.remove();
    if (needsCryptoStub) vi.unstubAllGlobals();
  });
});
