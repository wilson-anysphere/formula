import { describe, expect, test } from "vitest";

import { LayoutController } from "../layoutController.js";
import { LayoutWorkspaceManager, MemoryStorage } from "../layoutPersistence.js";
import { MAX_GRID_ZOOM } from "@formula/grid/node";

describe("LayoutController persistence", () => {
  test("setSplitPaneScroll can be applied without persisting, then flushed with persistNow()", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-1";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });

    // Persist an initial value so we can assert "no storage writes" on ephemeral updates.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();

    controller.setSplitPaneScroll("secondary", { scrollX: 1, scrollY: 2 }, { persist: false });

    expect(storage.getItem(key)).toBe(persistedBefore);
    expect(changeCount).toBe(2);

    controller.persistNow();

    const after = storage.getItem(key);
    expect(after).not.toBe(persistedBefore);
    expect(after).not.toBeNull();

    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(1);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(2);

    // persistNow should not emit an additional change event.
    expect(changeCount).toBe(2);
  });

  test("ephemeral updates can be applied without emitting a change event (emit:false)", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-emit-false";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });

    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();
    expect(changeCount).toBe(1);

    controller.setSplitPaneScroll("secondary", { scrollX: 2, scrollY: 3 }, { persist: false, emit: false });
    expect(changeCount).toBe(1);
    expect(storage.getItem(key)).toBe(persistedBefore);
    expect(controller.layout.splitView.panes.secondary.scrollX).toBe(2);
    expect(controller.layout.splitView.panes.secondary.scrollY).toBe(3);

    controller.persistNow();
    expect(changeCount).toBe(1);
    expect(storage.getItem(key)).not.toBe(persistedBefore);
  });

  test("setSplitRatio can be applied without persisting/emitting, then flushed with persistNow()", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-split-ratio";
    const controller = new LayoutController({ workbookId, workspaceManager });
    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;
 
    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });
 
    // Seed an initial persisted state.
    controller.setSplitDirection("vertical");
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();
    expect(changeCount).toBe(1);
 
    controller.setSplitRatio(0.05, { persist: false, emit: false });
    // Should clamp to [0.1, 0.9] even for silent updates.
    expect(controller.layout.splitView.ratio).toBeCloseTo(0.1, 5);
    // Silent + non-persisted updates should not write storage or emit.
    expect(storage.getItem(key)).toBe(persistedBefore);
    expect(changeCount).toBe(1);
 
    controller.persistNow();
    expect(changeCount).toBe(1);
    const after = storage.getItem(key);
    expect(after).not.toBe(persistedBefore);
    expect(after).not.toBeNull();
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.ratio).toBeCloseTo(0.1, 5);
  });

  test("setSplitRatio emit:false ignores invalid ratios (NaN/Infinity)", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-split-ratio-invalid";
    const controller = new LayoutController({ workbookId, workspaceManager });
    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });

    controller.setSplitDirection("vertical");
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();
    expect(changeCount).toBe(1);

    const ratioBefore = controller.layout.splitView.ratio;

    controller.setSplitRatio(Number.NaN, { persist: false, emit: false });
    controller.setSplitRatio(Number.POSITIVE_INFINITY, { persist: false, emit: false });
    controller.setSplitRatio(Number.NEGATIVE_INFINITY, { persist: false, emit: false });

    expect(controller.layout.splitView.ratio).toBe(ratioBefore);
    expect(storage.getItem(key)).toBe(persistedBefore);
    expect(changeCount).toBe(1);
  });

  test("setSplitPaneZoom can be applied without persisting/emitting, then flushed with persistNow()", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-split-zoom";
    const controller = new LayoutController({ workbookId, workspaceManager });
    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });

    // Seed initial persisted state.
    controller.setSplitDirection("vertical");
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();
    expect(changeCount).toBe(1);

    controller.setSplitPaneZoom("secondary", 10, { persist: false, emit: false });
    // Should clamp to the grid zoom bounds even for silent updates.
    expect(controller.layout.splitView.panes.secondary.zoom).toBe(MAX_GRID_ZOOM);
    expect(storage.getItem(key)).toBe(persistedBefore);
    expect(changeCount).toBe(1);

    controller.persistNow();
    expect(changeCount).toBe(1);
    const after = storage.getItem(key);
    expect(after).not.toBe(persistedBefore);
    expect(after).not.toBeNull();
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.zoom).toBe(MAX_GRID_ZOOM);
  });

  test("setSplitPaneScroll can be applied without persisting/emitting, then flushed with persistNow()", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-split-scroll";
    const controller = new LayoutController({ workbookId, workspaceManager });
    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    let changeCount = 0;
    controller.on("change", () => {
      changeCount += 1;
    });

    // Seed initial persisted state.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();
    expect(changeCount).toBe(1);

    controller.setSplitPaneScroll("secondary", { scrollX: 1e13, scrollY: -1e13 }, { persist: false, emit: false });
    // Should clamp to [-1e12, 1e12] even for silent updates.
    expect(controller.layout.splitView.panes.secondary.scrollX).toBe(1e12);
    expect(controller.layout.splitView.panes.secondary.scrollY).toBe(-1e12);
    expect(storage.getItem(key)).toBe(persistedBefore);
    expect(changeCount).toBe(1);

    controller.persistNow();
    expect(changeCount).toBe(1);
    const after = storage.getItem(key);
    expect(after).not.toBe(persistedBefore);
    expect(after).not.toBeNull();
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(1e12);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(-1e12);
  });

  test("switching workspaces persists pending ephemeral updates before reloading", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-2";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const key = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    // Seed initial persisted state.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(key);
    expect(persistedBefore).not.toBeNull();

    // Apply an ephemeral update (in-memory only).
    controller.setSplitPaneScroll("secondary", { scrollX: 5, scrollY: 6 }, { persist: false });
    expect(storage.getItem(key)).toBe(persistedBefore);

    // Switching workspaces reloads the layout; ensure we don't lose the in-memory change.
    controller.setWorkspace("analysis");

    const after = storage.getItem(key);
    expect(after).not.toBeNull();
    expect(after).not.toBe(persistedBefore);
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(5);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(6);
  });

  test("saveWorkspace(makeActive) persists pending ephemeral updates in the previous workspace", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-3";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const defaultKey = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    // Seed initial persisted state in the default workspace.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(defaultKey);
    expect(persistedBefore).not.toBeNull();

    // Apply an ephemeral update to the default workspace.
    controller.setSplitPaneScroll("secondary", { scrollX: 7, scrollY: 8 }, { persist: false });
    expect(storage.getItem(defaultKey)).toBe(persistedBefore);

    // Saving a new workspace and making it active should flush the pending update for the default workspace
    // (otherwise the default workspace would lose the latest scroll/zoom state when we switch away).
    controller.saveWorkspace("analysis", { makeActive: true });

    const after = storage.getItem(defaultKey);
    expect(after).not.toBeNull();
    expect(after).not.toBe(persistedBefore);
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(7);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(8);
  });

  test("deleteWorkspace for a different workspace persists pending ephemeral updates before reload", () => {
    const storage = new MemoryStorage();
    const keyPrefix = "test.layout";
    const workspaceManager = new LayoutWorkspaceManager({ storage, keyPrefix });
    const workbookId = "workbook-4";
    const controller = new LayoutController({ workbookId, workspaceManager });

    const defaultKey = `${keyPrefix}.workbook.${encodeURIComponent(workbookId)}.v1`;

    // Seed initial persisted state in the default workspace.
    controller.setSplitPaneScroll("secondary", { scrollX: 0, scrollY: 0 });
    const persistedBefore = storage.getItem(defaultKey);
    expect(persistedBefore).not.toBeNull();

    // Create a second workspace we can delete (without switching away from default).
    controller.saveWorkspace("analysis");

    // Apply an ephemeral update to the default workspace.
    controller.setSplitPaneScroll("secondary", { scrollX: 9, scrollY: 10 }, { persist: false });
    expect(storage.getItem(defaultKey)).toBe(persistedBefore);

    // Deleting a non-active workspace triggers a reload; ensure we flush pending updates first.
    controller.deleteWorkspace("analysis");

    const after = storage.getItem(defaultKey);
    expect(after).not.toBeNull();
    expect(after).not.toBe(persistedBefore);
    const parsed = JSON.parse(after!);
    expect(parsed.splitView.panes.secondary.scrollX).toBe(9);
    expect(parsed.splitView.panes.secondary.scrollY).toBe(10);
  });
});
