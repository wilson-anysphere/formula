import { describe, expect, it, vi } from "vitest";

import { MemoryStorage } from "../layout/layoutPersistence.js";
import {
  CONTRIBUTED_PANELS_SEED_STORE_KEY,
  readContributedPanelsSeedStore,
  removeSeedPanelsForExtension,
  setSeedPanelsForExtension,
} from "./contributedPanelsSeedStore.js";

describe("contributedPanelsSeedStore", () => {
  it("writes and reads panels for an extension", () => {
    const storage = new MemoryStorage();

    const ok = setSeedPanelsForExtension(storage as any, "acme.foo", [
      { id: "acme.foo.panel", title: "Foo Panel", icon: "foo-icon" },
    ]);

    expect(ok).toBe(true);
    expect(readContributedPanelsSeedStore(storage as any)).toEqual({
      "acme.foo.panel": { extensionId: "acme.foo", title: "Foo Panel", icon: "foo-icon" },
    });
  });

  it("replaces contributed panel metadata for an extension on update", () => {
    const storage = new MemoryStorage();

    expect(
      setSeedPanelsForExtension(storage as any, "acme.foo", [
        { id: "acme.foo.panel1", title: "Panel 1" },
        { id: "acme.foo.panel2", title: "Panel 2" },
      ]),
    ).toBe(true);

    expect(
      setSeedPanelsForExtension(storage as any, "acme.foo", [{ id: "acme.foo.panel2", title: "Panel 2 (new)" }]),
    ).toBe(true);

    const data = readContributedPanelsSeedStore(storage as any);
    expect(Object.keys(data).sort()).toEqual(["acme.foo.panel2"]);
    expect(data["acme.foo.panel2"]).toMatchObject({ extensionId: "acme.foo", title: "Panel 2 (new)" });
  });

  it("removes contributed panels for an extension on uninstall", () => {
    const storage = new MemoryStorage();

    setSeedPanelsForExtension(storage as any, "acme.foo", [{ id: "acme.foo.panel", title: "Foo Panel" }]);
    setSeedPanelsForExtension(storage as any, "acme.bar", [{ id: "acme.bar.panel", title: "Bar Panel" }]);

    removeSeedPanelsForExtension(storage as any, "acme.foo");

    expect(readContributedPanelsSeedStore(storage as any)).toEqual({
      "acme.bar.panel": { extensionId: "acme.bar", title: "Bar Panel" },
    });
  });

  it("removes the seed store key when the last contributed panel is removed", () => {
    const storage = new MemoryStorage();

    setSeedPanelsForExtension(storage as any, "acme.foo", [{ id: "acme.foo.panel", title: "Foo Panel" }]);
    expect(storage.getItem(CONTRIBUTED_PANELS_SEED_STORE_KEY)).not.toBeNull();

    removeSeedPanelsForExtension(storage as any, "acme.foo");

    expect(readContributedPanelsSeedStore(storage as any)).toEqual({});
    expect(storage.getItem(CONTRIBUTED_PANELS_SEED_STORE_KEY)).toBeNull();
  });

  it("rejects conflicting panel ids across extensions without mutating the store", () => {
    const storage = new MemoryStorage();

    setSeedPanelsForExtension(storage as any, "acme.foo", [{ id: "shared.panel", title: "Foo Panel" }]);

    const onError = vi.fn();
    const ok = setSeedPanelsForExtension(
      storage as any,
      "acme.bar",
      [{ id: "shared.panel", title: "Bar Panel" }],
      { onError },
    );

    expect(ok).toBe(false);
    expect(onError).toHaveBeenCalled();

    const data = readContributedPanelsSeedStore(storage as any);
    expect(data["shared.panel"]).toMatchObject({ extensionId: "acme.foo", title: "Foo Panel" });
  });
});
