import assert from "node:assert/strict";
import test from "node:test";

import { DesktopSessionController } from "../src/layout/sessionController.js";
import { LayoutWorkspaceManager, MemoryStorage } from "../src/layout/layoutPersistence.js";
import { WindowingController } from "../src/layout/windowingController.js";
import { WindowingSessionManager } from "../src/layout/windowingPersistence.js";
import { getPanelPlacement } from "../src/layout/layoutState.js";
import { PANEL_REGISTRY, PanelIds } from "../src/panels/panelRegistry.js";

test("DesktopSessionController wires windowing + per-window layout controllers", () => {
  const storage = new MemoryStorage();

  const layoutWorkspaceManager = new LayoutWorkspaceManager({ storage, panelRegistry: PANEL_REGISTRY });
  const windowingSessionManager = new WindowingSessionManager({ storage });
  const windowingController = new WindowingController({ sessionManager: windowingSessionManager });

  const session1 = new DesktopSessionController({ layoutWorkspaceManager, windowingController });

  const winDefault = session1.openWorkbookWindow("workbook-a", { workspaceId: "default" });
  const winAnalysis = session1.openWorkbookWindow("workbook-a", { workspaceId: "analysis" });

  const layoutDefault = session1.getLayoutController(winDefault);
  const layoutAnalysis = session1.getLayoutController(winAnalysis);

  assert.ok(layoutDefault);
  assert.ok(layoutAnalysis);

  layoutDefault.openPanel(PanelIds.AI_CHAT);
  layoutDefault.dockPanel(PanelIds.AI_CHAT, "left");

  layoutAnalysis.openPanel(PanelIds.VERSION_HISTORY);
  layoutAnalysis.dockPanel(PanelIds.VERSION_HISTORY, "right");

  // "Reload app": new windowing controller (loads persisted window session) + new session controller.
  const windowingController2 = new WindowingController({ sessionManager: windowingSessionManager });
  const session2 = new DesktopSessionController({ layoutWorkspaceManager, windowingController: windowingController2 });

  const restoredDefault = session2.getLayoutController(winDefault);
  const restoredAnalysis = session2.getLayoutController(winAnalysis);

  assert.ok(restoredDefault);
  assert.ok(restoredAnalysis);

  assert.deepEqual(getPanelPlacement(restoredDefault.layout, PanelIds.AI_CHAT), { kind: "docked", side: "left" });
  assert.deepEqual(getPanelPlacement(restoredDefault.layout, PanelIds.VERSION_HISTORY), { kind: "closed" });

  assert.deepEqual(getPanelPlacement(restoredAnalysis.layout, PanelIds.VERSION_HISTORY), { kind: "docked", side: "right" });
});

