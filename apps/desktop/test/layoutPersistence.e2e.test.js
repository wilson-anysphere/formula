import assert from "node:assert/strict";
import test from "node:test";

import { LayoutWorkspaceManager, MemoryStorage } from "../src/layout/layoutPersistence.js";
import { dockPanel, getPanelPlacement, openPanel } from "../src/layout/layoutState.js";
import { PANEL_REGISTRY, PanelIds } from "../src/panels/panelRegistry.js";

test("E2E: docked panel layout persists per-workbook and restores on reload", () => {
  const storage = new MemoryStorage();

  const manager1 = new LayoutWorkspaceManager({ storage, panelRegistry: PANEL_REGISTRY });
  const workbookId = "workbook-123";

  let layout = manager1.loadWorkbookLayout(workbookId, { primarySheetId: "Sheet1" });

  layout = openPanel(layout, PanelIds.AI_CHAT, { panelRegistry: PANEL_REGISTRY });
  layout = dockPanel(layout, PanelIds.AI_CHAT, "left");

  manager1.saveWorkbookLayout(workbookId, layout);

  const manager2 = new LayoutWorkspaceManager({ storage, panelRegistry: PANEL_REGISTRY });
  const restored = manager2.loadWorkbookLayout(workbookId, { primarySheetId: "Sheet1" });

  assert.deepEqual(getPanelPlacement(restored, PanelIds.AI_CHAT), { kind: "docked", side: "left" });
  assert.equal(restored.docks.left.active, PanelIds.AI_CHAT);
});

test("LayoutWorkspaceManager falls back to global default layout when workbook has none", () => {
  const storage = new MemoryStorage();
  const workbookId = "workbook-456";

  const manager = new LayoutWorkspaceManager({ storage, panelRegistry: PANEL_REGISTRY });

  let globalLayout = manager.loadWorkbookLayout("temp", { primarySheetId: "Sheet1" });
  globalLayout = openPanel(globalLayout, PanelIds.VERSION_HISTORY, { panelRegistry: PANEL_REGISTRY });
  globalLayout = dockPanel(globalLayout, PanelIds.VERSION_HISTORY, "left");

  manager.saveGlobalDefaultLayout(globalLayout);

  const fromWorkbook = manager.loadWorkbookLayout(workbookId, { primarySheetId: "Sheet1" });
  assert.deepEqual(getPanelPlacement(fromWorkbook, PanelIds.VERSION_HISTORY), { kind: "docked", side: "left" });

  // Workbook override should win over global.
  let workbookLayout = openPanel(fromWorkbook, PanelIds.AI_CHAT, { panelRegistry: PANEL_REGISTRY });
  workbookLayout = dockPanel(workbookLayout, PanelIds.AI_CHAT, "right");
  manager.saveWorkbookLayout(workbookId, workbookLayout);

  const overridden = manager.loadWorkbookLayout(workbookId, { primarySheetId: "Sheet1" });
  assert.deepEqual(getPanelPlacement(overridden, PanelIds.AI_CHAT), { kind: "docked", side: "right" });
  assert.deepEqual(getPanelPlacement(overridden, PanelIds.VERSION_HISTORY), { kind: "docked", side: "left" });
});

test("LayoutWorkspaceManager supports multiple named workspaces per workbook", () => {
  const storage = new MemoryStorage();
  const workbookId = "workbook-789";

  const manager = new LayoutWorkspaceManager({ storage, panelRegistry: PANEL_REGISTRY });

  let defaultLayout = manager.loadWorkbookLayout(workbookId, { primarySheetId: "Sheet1" });
  defaultLayout = openPanel(defaultLayout, PanelIds.AI_CHAT, { panelRegistry: PANEL_REGISTRY });
  defaultLayout = dockPanel(defaultLayout, PanelIds.AI_CHAT, "left");
  manager.saveWorkbookLayout(workbookId, defaultLayout);

  let analysisLayout = openPanel(defaultLayout, PanelIds.VERSION_HISTORY, { panelRegistry: PANEL_REGISTRY });
  analysisLayout = dockPanel(analysisLayout, PanelIds.VERSION_HISTORY, "right");
  manager.saveWorkbookWorkspace(workbookId, "analysis", { name: "Analysis", layout: analysisLayout, makeActive: true });

  const fromAnalysis = manager.loadWorkbookLayout(workbookId, { primarySheetId: "Sheet1" });
  assert.deepEqual(getPanelPlacement(fromAnalysis, PanelIds.VERSION_HISTORY), { kind: "docked", side: "right" });

  manager.setActiveWorkbookWorkspace(workbookId, "default");
  const fromDefault = manager.loadWorkbookLayout(workbookId, { primarySheetId: "Sheet1" });
  assert.deepEqual(getPanelPlacement(fromDefault, PanelIds.AI_CHAT), { kind: "docked", side: "left" });
  assert.deepEqual(getPanelPlacement(fromDefault, PanelIds.VERSION_HISTORY), { kind: "closed" });

  const workspaces = manager.listWorkbookWorkspaces(workbookId);
  assert.deepEqual(
    workspaces.map((w) => ({ id: w.id, active: w.active })),
    [
      { id: "default", active: true },
      { id: "analysis", active: false },
    ],
  );

  manager.deleteWorkbookWorkspace(workbookId, "analysis");
  assert.deepEqual(
    manager.listWorkbookWorkspaces(workbookId).map((w) => w.id),
    ["default"],
  );
});
