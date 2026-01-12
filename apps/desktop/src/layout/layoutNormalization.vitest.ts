import { describe, expect, it } from "vitest";

import { deserializeLayout } from "./layoutSerializer.js";
import { PanelRegistry } from "../panels/panelRegistry.js";
import { MemoryStorage } from "./layoutPersistence.js";
import { seedPanelRegistryFromContributedPanelsSeedStore, writeContributedPanelsSeedStore } from "../extensions/contributedPanelsSeedStore.js";

describe("layout normalization - extension panels", () => {
  it("drops unknown extension panel ids when panelRegistry is not seeded", () => {
    const panelId = "acme.foo.panel";
    const layoutJson = JSON.stringify({
      version: 1,
      docks: {
        right: {
          panels: [panelId],
        },
      },
    });

    const panelRegistry = new PanelRegistry();
    const layout = deserializeLayout(layoutJson, { panelRegistry });

    expect(layout.docks.right.panels).not.toContain(panelId);
  });

  it("retains extension panel ids when the seed store is applied before deserialization", () => {
    const panelId = "acme.foo.panel";
    const extensionId = "acme.foo";

    const storage = new MemoryStorage();
    writeContributedPanelsSeedStore(storage as any, {
      [panelId]: {
        extensionId,
        title: "Acme Panel",
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
          panels: [panelId],
        },
      },
    });

    const layout = deserializeLayout(layoutJson, { panelRegistry });

    expect(layout.docks.right.panels).toContain(panelId);
  });
});

