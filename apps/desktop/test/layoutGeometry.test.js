import assert from "node:assert/strict";
import test from "node:test";

import { computeWorkspaceRects } from "../src/layout/layoutGeometry.js";
import {
  createDefaultLayout,
  dockPanel,
  floatPanel,
  getPanelPlacement,
  setDockSize,
  setSplitDirection,
  snapFloatingPanel,
} from "../src/layout/layoutState.js";
import { PanelIds } from "../src/panels/panelRegistry.js";

test("snapFloatingPanel only docks when within threshold distance", () => {
  const viewport = { width: 800, height: 600 };

  let layout = createDefaultLayout({ primarySheetId: "Sheet1" });
  layout = floatPanel(layout, PanelIds.AI_CHAT, { x: -500, y: 50, width: 300, height: 300 });

  const snapped = snapFloatingPanel(layout, PanelIds.AI_CHAT, viewport, { threshold: 24 });
  assert.deepEqual(getPanelPlacement(snapped, PanelIds.AI_CHAT), { kind: "floating" });

  const nearEdge = floatPanel(layout, PanelIds.AI_CHAT, { x: 10, y: 50, width: 300, height: 300 });
  const snappedNear = snapFloatingPanel(nearEdge, PanelIds.AI_CHAT, viewport, { threshold: 24 });
  assert.deepEqual(getPanelPlacement(snappedNear, PanelIds.AI_CHAT), { kind: "docked", side: "left" });
});

test("computeWorkspaceRects accounts for dock zones + split panes", () => {
  let layout = createDefaultLayout({ primarySheetId: "Sheet1" });
  layout = dockPanel(layout, PanelIds.AI_CHAT, "left");
  layout = setDockSize(layout, "left", 200);
  layout = setSplitDirection(layout, "vertical", 0.5);

  const rects = computeWorkspaceRects(layout, { width: 1000, height: 800 }, { gutter: 4 });

  assert.deepEqual(rects.content, { x: 200, y: 0, width: 800, height: 800 });
  assert.deepEqual(rects.docks.left, { x: 0, y: 0, width: 200, height: 800 });

  assert.equal(rects.split.direction, "vertical");
  assert.deepEqual(rects.split.primary, { x: 200, y: 0, width: 398, height: 800 });
  assert.deepEqual(rects.split.gutter, { x: 598, y: 0, width: 4, height: 800 });
  assert.deepEqual(rects.split.secondary, { x: 602, y: 0, width: 398, height: 800 });
});

