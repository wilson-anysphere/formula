import assert from "node:assert/strict";
import test from "node:test";

import { LayoutController } from "../src/layout/layoutController.js";
import { LayoutWorkspaceManager, MemoryStorage } from "../src/layout/layoutPersistence.js";
import { getPanelPlacement } from "../src/layout/layoutState.js";
import { PANEL_REGISTRY, PanelIds } from "../src/panels/panelRegistry.js";

test("LayoutController persists layout changes and restores on reload", () => {
  const storage = new MemoryStorage();
  const workspaceManager = new LayoutWorkspaceManager({ storage, panelRegistry: PANEL_REGISTRY });
  const workbookId = "workbook-controller";

  const controller1 = new LayoutController({ workbookId, workspaceManager, primarySheetId: "Sheet1" });

  controller1.openPanel(PanelIds.AI_CHAT);
  controller1.dockPanel(PanelIds.AI_CHAT, "left");

  const controller2 = new LayoutController({ workbookId, workspaceManager, primarySheetId: "Sheet1" });

  assert.deepEqual(getPanelPlacement(controller2.layout, PanelIds.AI_CHAT), { kind: "docked", side: "left" });
});

test("LayoutController can switch between named workspaces", () => {
  const storage = new MemoryStorage();
  const workspaceManager = new LayoutWorkspaceManager({ storage, panelRegistry: PANEL_REGISTRY });
  const workbookId = "workbook-workspaces-controller";

  const controller = new LayoutController({ workbookId, workspaceManager, primarySheetId: "Sheet1" });

  controller.openPanel(PanelIds.AI_CHAT);
  controller.dockPanel(PanelIds.AI_CHAT, "left");
  controller.saveWorkspace("analysis", { name: "Analysis", makeActive: true });

  // In analysis workspace, AI chat should still be docked left.
  assert.deepEqual(getPanelPlacement(controller.layout, PanelIds.AI_CHAT), { kind: "docked", side: "left" });

  // Now modify analysis layout.
  controller.openPanel(PanelIds.VERSION_HISTORY);
  controller.dockPanel(PanelIds.VERSION_HISTORY, "right");

  // Switch back to default: version history should not be present.
  controller.setActiveWorkspace("default");
  assert.deepEqual(getPanelPlacement(controller.layout, PanelIds.VERSION_HISTORY), { kind: "closed" });

  // Switch to analysis: version history should come back.
  controller.setActiveWorkspace("analysis");
  assert.deepEqual(getPanelPlacement(controller.layout, PanelIds.VERSION_HISTORY), { kind: "docked", side: "right" });
});

