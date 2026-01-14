// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";

import { ConflictUiController } from "./conflict-ui-controller.js";

describe("ConflictUiController", () => {
  it("renders a Jump to cell button and invokes the callback with the conflict cell", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const onNavigateToCell = vi.fn();

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
      onNavigateToCell,
    });

    ui.addConflict({
      id: "c1",
      kind: "formula",
      cell: { sheetId: "Sheet1", row: 3, col: 2 },
      cellKey: "Sheet1:3:2",
      localFormula: "=1",
      remoteFormula: "=2",
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const jump = container.querySelector<HTMLButtonElement>('[data-testid="conflict-jump-to-cell"]');
    expect(jump).not.toBeNull();
    jump!.click();

    expect(onNavigateToCell).toHaveBeenCalledTimes(1);
    expect(onNavigateToCell).toHaveBeenCalledWith({ sheetId: "Sheet1", row: 3, col: 2 });

    ui.destroy();
    container.remove();
  });

  it("trims sheetId before invoking onNavigateToCell", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);
    const onNavigateToCell = vi.fn();

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
      onNavigateToCell,
    });

    ui.addConflict({
      id: "c_trim_sheet",
      kind: "formula",
      cell: { sheetId: "  Sheet1  ", row: 3, col: 2 },
      cellKey: "Sheet1:3:2",
      localFormula: "=1",
      remoteFormula: "=2",
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const jump = container.querySelector<HTMLButtonElement>('[data-testid="conflict-jump-to-cell"]');
    expect(jump).not.toBeNull();
    jump!.click();

    expect(onNavigateToCell).toHaveBeenCalledWith({ sheetId: "Sheet1", row: 3, col: 2 });

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

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
      onNavigateToCell,
    });

    ui.addConflict({
      id: "c_jump_throw",
      kind: "formula",
      cell: { sheetId: "Sheet1", row: 3, col: 2 },
      cellKey: "Sheet1:3:2",
      localFormula: "=1",
      remoteFormula: "=2",
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const jump = container.querySelector<HTMLButtonElement>('[data-testid="conflict-jump-to-cell"]');
    expect(jump).not.toBeNull();
    expect(() => jump!.click()).not.toThrow();
    expect(onNavigateToCell).toHaveBeenCalledWith({ sheetId: "Sheet1", row: 3, col: 2 });

    ui.destroy();
    container.remove();
  });

  it("applies a user label resolver when rendering the remote (theirs) panel label", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
      resolveUserLabel: (userId: string) => (userId === "u2" ? "  Bob  " : userId),
    });

    ui.addConflict({
      id: "c2",
      kind: "formula",
      cell: { sheetId: "Sheet1", row: 0, col: 0 },
      cellKey: "Sheet1:0:0",
      localFormula: "=1",
      remoteFormula: "=2",
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const remotePanel = container.querySelector<HTMLElement>('[data-testid="conflict-remote"]');
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

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
      resolveUserLabel: () => {
        throw new Error("boom");
      },
    });

    ui.addConflict({
      id: "c3",
      kind: "formula",
      cell: { sheetId: "Sheet1", row: 0, col: 0 },
      cellKey: "Sheet1:0:0",
      localFormula: "=1",
      remoteFormula: "=2",
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const remotePanel = container.querySelector<HTMLElement>('[data-testid="conflict-remote"]');
    expect(remotePanel).not.toBeNull();

    const label = remotePanel!.querySelector<HTMLElement>(".conflict-dialog__panel-label");
    expect(label?.textContent).toBe("Theirs (u2)");

    ui.destroy();
    container.remove();
  });

  it("renders a token diff view for formula conflicts", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "c4",
      kind: "formula",
      cell: { sheetId: "Sheet1", row: 0, col: 0 },
      cellKey: "Sheet1:0:0",
      localFormula: "=A1",
      remoteFormula: "=A2",
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const diff = container.querySelector<HTMLElement>('[data-testid="conflict-formula-diff"]');
    expect(diff).not.toBeNull();
    expect(diff!.querySelector(".formula-diff-op--delete")).not.toBeNull();
    expect(diff!.querySelector(".formula-diff-op--insert")).not.toBeNull();

    ui.destroy();
    container.remove();
  });

  it("renders a token diff view for content conflicts when both sides are formulas", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "c5",
      kind: "content",
      cell: { sheetId: "Sheet1", row: 0, col: 0 },
      cellKey: "Sheet1:0:0",
      local: { type: "formula", formula: "=A1" },
      remote: { type: "formula", formula: "=A2" },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const diff = container.querySelector<HTMLElement>('[data-testid="conflict-formula-diff"]');
    expect(diff).not.toBeNull();
    expect(diff!.querySelector(".formula-diff-op--delete")).not.toBeNull();
    expect(diff!.querySelector(".formula-diff-op--insert")).not.toBeNull();

    ui.destroy();
    container.remove();
  });

  it("renders a token diff view for content conflicts when one side is a formula", () => {
    const container = document.createElement("div");
    document.body.appendChild(container);

    const resolveConflict = vi.fn(() => true);

    const ui = new ConflictUiController({
      container,
      monitor: { resolveConflict },
    });

    ui.addConflict({
      id: "c6",
      kind: "content",
      cell: { sheetId: "Sheet1", row: 0, col: 0 },
      cellKey: "Sheet1:0:0",
      local: { type: "value", value: 123 },
      remote: { type: "formula", formula: "=A1" },
      remoteUserId: "u2",
      detectedAt: 0,
    });

    container.querySelector<HTMLButtonElement>('[data-testid="conflict-toast-open"]')?.click();

    const diff = container.querySelector<HTMLElement>('[data-testid="conflict-formula-diff"]');
    expect(diff).not.toBeNull();
    expect(diff!.querySelector(".formula-diff-op--insert")).not.toBeNull();

    ui.destroy();
    container.remove();
  });
});
