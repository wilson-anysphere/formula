import { describe, expect, it } from "vitest";

import { deserializeLayout } from "./layoutSerializer.js";
import { PanelRegistry } from "../panels/panelRegistry.js";
import { MemoryStorage } from "./layoutPersistence.js";
import { seedPanelRegistryFromContributedPanelsSeedStore, writeContributedPanelsSeedStore } from "../extensions/contributedPanelsSeedStore.js";

describe("layout normalization - extension panels", () => {
  it("drops unknown extension panel ids when panelRegistry is not seeded", () => {
    const dockedPanelId = "acme.foo.panel";
    const floatingPanelId = "acme.foo.floatingPanel";
    const layoutJson = JSON.stringify({
      version: 1,
      docks: {
        right: {
          panels: [dockedPanelId],
        },
      },
      floating: {
        [floatingPanelId]: { x: 10, y: 20, width: 300, height: 200, minimized: false },
      },
    });

    const panelRegistry = new PanelRegistry();
    const layout = deserializeLayout(layoutJson, { panelRegistry });

    expect(layout.docks.right.panels).not.toContain(dockedPanelId);
    expect(layout.floating).not.toHaveProperty(floatingPanelId);
  });

  it("retains extension panel ids when the seed store is applied before deserialization", () => {
    const dockedPanelId = "acme.foo.panel";
    const floatingPanelId = "acme.foo.floatingPanel";
    const extensionId = "acme.foo";

    const storage = new MemoryStorage();
    writeContributedPanelsSeedStore(storage as any, {
      [dockedPanelId]: {
        extensionId,
        title: "Acme Panel",
        icon: null,
        defaultDock: "right",
      },
      [floatingPanelId]: {
        extensionId,
        title: "Acme Floating Panel",
        icon: null,
        defaultDock: "right",
      },
    });

    const panelRegistry = new PanelRegistry();
    seedPanelRegistryFromContributedPanelsSeedStore(storage as any, panelRegistry);

    const layoutJson = JSON.stringify({
      version: 1,
      docks: {
        right: {
          panels: [dockedPanelId],
        },
      },
      floating: {
        [floatingPanelId]: { x: 10, y: 20, width: 300, height: 200, minimized: false },
      },
    });

    const layout = deserializeLayout(layoutJson, { panelRegistry });

    expect(layout.docks.right.panels).toContain(dockedPanelId);
    expect(layout.floating).toHaveProperty(floatingPanelId);
  });
});
