// @vitest-environment jsdom
import { afterEach, describe, expect, it } from "vitest";

import { ThemeController } from "../themeController.js";
import { MEDIA } from "../systemPreferences.js";
import { createMemoryStorage, saveAppearanceSettings } from "../../settings/appearance/index.js";

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
  // Reset document state between tests to avoid leakage.
  document.documentElement.removeAttribute("data-theme");
  document.documentElement.removeAttribute("data-reduced-motion");
});

describe("ThemeController (jsdom)", () => {
  it("start() applies persisted theme preference from storage", () => {
    const storage = createMemoryStorage();
    saveAppearanceSettings({ themePreference: "dark" }, storage);

    const controller = new ThemeController({ document, storage });
    controller.start();

    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");
    controller.stop();
  });

  it("resolves system theme from matchMedia and updates when preferences change", () => {
    const env = createMatchMediaStub({
      [MEDIA.prefersDark]: false,
      [MEDIA.forcedColors]: false,
      [MEDIA.prefersContrastMore]: false,
    });

    const controller = new ThemeController({ document, env, storage: createMemoryStorage() });
    controller.start();

    expect(document.documentElement.getAttribute("data-theme")).toBe("light");

    env.setMatches(MEDIA.prefersDark, true);
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");

    // High contrast takes precedence over light/dark.
    env.setMatches(MEDIA.forcedColors, true);
    expect(document.documentElement.getAttribute("data-theme")).toBe("high-contrast");

    env.setMatches(MEDIA.forcedColors, false);
    expect(document.documentElement.getAttribute("data-theme")).toBe("dark");

    controller.stop();
  });

  it("applies reduced motion attribute from prefers-reduced-motion", () => {
    const env = createMatchMediaStub({
      [MEDIA.reducedMotion]: false,
    });

    const controller = new ThemeController({ document, env, storage: createMemoryStorage() });
    controller.start();

    expect(document.documentElement.getAttribute("data-reduced-motion")).toBe("false");

    env.setMatches(MEDIA.reducedMotion, true);
    expect(document.documentElement.getAttribute("data-reduced-motion")).toBe("true");

    controller.stop();
  });
});
