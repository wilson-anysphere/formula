// @vitest-environment jsdom
import { afterEach, describe, expect, it } from "vitest";

import { CommandRegistry } from "../extensions/commandRegistry.js";
import { ThemeController } from "../theme/themeController.js";
import { MEDIA } from "../theme/systemPreferences.js";
import { createMemoryStorage, getThemePreference, saveAppearanceSettings } from "../settings/appearance/index.js";

import { registerBuiltinCommands } from "./registerBuiltinCommands.js";

type MatchMediaStub = {
  matchMedia: (query: string) => MediaQueryList;
  setMatches: (query: string, matches: boolean) => void;
};

function createMatchMediaStub(initialMatches: Record<string, boolean> = {}): MatchMediaStub {
  const listeners = new Map<string, Set<(event: { matches: boolean }) => void>>();
  const matchesByQuery: Record<string, boolean> = { ...initialMatches };

  function getListeners(query: string) {
    const existing = listeners.get(query);
    if (existing) return existing;
    const set = new Set<(event: { matches: boolean }) => void>();
    listeners.set(query, set);
    return set;
  }

  function notify(query: string) {
    for (const cb of getListeners(query)) cb({ matches: Boolean(matchesByQuery[query]) });
  }

  return {
    matchMedia(query: string): MediaQueryList {
      // Provide both modern (addEventListener) and legacy (addListener) APIs since
      // `subscribeToMediaQuery` supports both.
      const api: MediaQueryList = {
        media: query,
        get matches() {
          return Boolean(matchesByQuery[query]);
        },
        onchange: null,
        addEventListener(_type: string, cb: EventListenerOrEventListenerObject) {
          getListeners(query).add(cb as unknown as (event: { matches: boolean }) => void);
        },
        removeEventListener(_type: string, cb: EventListenerOrEventListenerObject) {
          getListeners(query).delete(cb as unknown as (event: { matches: boolean }) => void);
        },
        addListener(cb) {
          getListeners(query).add(cb as unknown as (event: { matches: boolean }) => void);
        },
        removeListener(cb) {
          getListeners(query).delete(cb as unknown as (event: { matches: boolean }) => void);
        },
        dispatchEvent(_event) {
          return true;
        },
      };
      return api;
    },
    setMatches(query: string, value: boolean) {
      matchesByQuery[query] = Boolean(value);
      notify(query);
    },
  };
}

afterEach(() => {
  document.documentElement.removeAttribute("data-theme");
  document.documentElement.removeAttribute("data-reduced-motion");
});

describe("theme toggle commands", () => {
  it("executing view.theme.dark updates data-theme and persists storage", async () => {
    const storage = createMemoryStorage();
    const controller = new ThemeController({ document, storage });
    controller.start();

    const commandRegistry = new CommandRegistry();
    registerBuiltinCommands({ commandRegistry, app: {} as any, layoutController: {} as any, themeController: controller });

    expect(commandRegistry.getCommand("view.theme.dark")).toBeDefined();

    expect(document.documentElement.getAttribute("data-theme")).toBe("light");
    expect(getThemePreference(storage)).toBe("system");

    await commandRegistry.executeCommand("view.theme.dark");
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    expect(getThemePreference(storage)).toBe("dark");

    // Simulate reload: a new ThemeController should pick up the persisted preference.
    controller.stop();
    document.documentElement.removeAttribute("data-theme");

    const controller2 = new ThemeController({ document, storage });
    controller2.start();
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    controller2.stop();
  });

  it("executing view.theme.system makes ThemeController follow matchMedia", async () => {
    const storage = createMemoryStorage();
    saveAppearanceSettings({ themePreference: "dark" }, storage);

    const env = createMatchMediaStub({
      [MEDIA.prefersDark]: false,
      [MEDIA.forcedColors]: false,
      [MEDIA.prefersContrastMore]: false,
      [MEDIA.reducedMotion]: false,
    });

    const controller = new ThemeController({ document, storage, env });
    controller.start();

    const commandRegistry = new CommandRegistry();
    registerBuiltinCommands({ commandRegistry, app: {} as any, layoutController: {} as any, themeController: controller });

    // Storage forces a fixed theme; the controller should not react to system changes.
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    env.setMatches(MEDIA.prefersDark, true);
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");

    // Switch to system mode; should now resolve from matchMedia and subscribe to changes.
    env.setMatches(MEDIA.prefersDark, false);
    await commandRegistry.executeCommand("view.theme.system");
    expect(getThemePreference(storage)).toBe("system");
    expect(document.documentElement.getAttribute("data-theme")).toBe("light");

    env.setMatches(MEDIA.prefersDark, true);
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    controller.stop();
  });
});
