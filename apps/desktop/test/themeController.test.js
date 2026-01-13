import test from "node:test";
import assert from "node:assert/strict";

import { ThemeController, resolveTheme } from "../src/theme/index.js";
import {
  createMemoryStorage,
  loadAppearanceSettings,
  saveAppearanceSettings,
} from "../src/settings/appearance/index.js";

function createMatchMediaStub(initialMatches = {}) {
  const listeners = new Map();
  const matches = { ...initialMatches };

  function getListeners(query) {
    if (!listeners.has(query)) listeners.set(query, new Set());
    return listeners.get(query);
  }

  function notify(query) {
    const set = getListeners(query);
    for (const cb of set) cb({ matches: Boolean(matches[query]) });
  }

  return {
    matchMedia(query) {
      const api = {
        media: query,
        get matches() {
          return Boolean(matches[query]);
        },
        addEventListener(_type, cb) {
          getListeners(query).add(cb);
        },
        removeEventListener(_type, cb) {
          getListeners(query).delete(cb);
        },
        addListener(cb) {
          getListeners(query).add(cb);
        },
        removeListener(cb) {
          getListeners(query).delete(cb);
        }
      };
      return api;
    },
    setMatches(query, value) {
      matches[query] = Boolean(value);
      notify(query);
    }
  };
}

function createStubDocument() {
  const attrs = new Map();
  return {
    documentElement: {
      style: {},
      setAttribute(name, value) {
        attrs.set(name, String(value));
      },
      getAttribute(name) {
        return attrs.has(name) ? attrs.get(name) : null;
      }
    }
  };
}

test("resolveTheme prefers high contrast when forced-colors is active", () => {
  const env = createMatchMediaStub({ "(forced-colors: active)": true });
  assert.equal(resolveTheme("system", env), "high-contrast");
});

test("loadAppearanceSettings defaults to light when the storage entry is missing/invalid", () => {
  const storage = createMemoryStorage();
  const key = "formula.settings.appearance.v1";

  // Missing key.
  assert.deepEqual(loadAppearanceSettings(storage), { themePreference: "light" });

  // Invalid JSON.
  storage.setItem(key, "not-json");
  assert.deepEqual(loadAppearanceSettings(storage), { themePreference: "light" });

  // Invalid theme value.
  storage.setItem(key, JSON.stringify({ themePreference: "banana" }));
  assert.deepEqual(loadAppearanceSettings(storage), { themePreference: "light" });
});

test("ThemeController applies resolved theme + reduced motion and persists preference", () => {
  const env = createMatchMediaStub({
    "(prefers-color-scheme: dark)": true,
    "(prefers-reduced-motion: reduce)": false
  });
  const document = createStubDocument();
  const storage = createMemoryStorage();

  const controller = new ThemeController({ env, document, storage });
  controller.start();

  // UX default: Light theme, regardless of the OS preferred color scheme.
  assert.equal(controller.getThemePreference(), "light");
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");
  assert.equal(document.documentElement.getAttribute("data-reduced-motion"), "false");

  controller.setThemePreference("dark");
  assert.equal(document.documentElement.getAttribute("data-theme"), "dark");

  // System changes should not affect a user override.
  env.setMatches("(prefers-color-scheme: dark)", false);
  assert.equal(document.documentElement.getAttribute("data-theme"), "dark");

  // Persistence.
  const doc2 = createStubDocument();
  const controller2 = new ThemeController({ env, document: doc2, storage });
  controller2.start();
  assert.equal(controller2.getThemePreference(), "dark");
  assert.equal(doc2.documentElement.getAttribute("data-theme"), "dark");
});

test("ThemeController updates theme when following system preferences", () => {
  const env = createMatchMediaStub({ "(prefers-color-scheme: dark)": false });
  const document = createStubDocument();
  const storage = createMemoryStorage();

  const controller = new ThemeController({ env, document, storage });
  controller.start();

  assert.equal(controller.getThemePreference(), "light");
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");

  controller.setThemePreference("system");
  assert.equal(controller.getThemePreference(), "system");

  env.setMatches("(prefers-color-scheme: dark)", true);
  assert.equal(document.documentElement.getAttribute("data-theme"), "dark");
});

test("ThemeController updates to high contrast when forced-colors toggles while following system", () => {
  const env = createMatchMediaStub({
    "(prefers-color-scheme: dark)": false,
    "(forced-colors: active)": false
  });
  const document = createStubDocument();
  const storage = createMemoryStorage();

  const controller = new ThemeController({ env, document, storage });
  controller.start();

  assert.equal(controller.getThemePreference(), "light");
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");

  controller.setThemePreference("system");
  assert.equal(controller.getThemePreference(), "system");

  env.setMatches("(forced-colors: active)", true);
  assert.equal(document.documentElement.getAttribute("data-theme"), "high-contrast");
});

test("ThemeController updates to high contrast when prefers-contrast toggles while following system", () => {
  const env = createMatchMediaStub({
    "(prefers-color-scheme: dark)": false,
    "(prefers-contrast: more)": false
  });
  const document = createStubDocument();
  const storage = createMemoryStorage();

  const controller = new ThemeController({ env, document, storage });
  controller.start();

  assert.equal(controller.getThemePreference(), "light");
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");

  controller.setThemePreference("system");
  assert.equal(controller.getThemePreference(), "system");

  env.setMatches("(prefers-contrast: more)", true);
  assert.equal(document.documentElement.getAttribute("data-theme"), "high-contrast");
});

test("ThemeController updates reduced motion attribute when preference changes", () => {
  const env = createMatchMediaStub({ "(prefers-reduced-motion: reduce)": false });
  const document = createStubDocument();
  const controller = new ThemeController({ env, document, storage: createMemoryStorage() });
  controller.start();

  assert.equal(document.documentElement.getAttribute("data-reduced-motion"), "false");
  env.setMatches("(prefers-reduced-motion: reduce)", true);
  assert.equal(document.documentElement.getAttribute("data-reduced-motion"), "true");
});

test("ThemeController prefers high contrast when forced-colors is active (system theme)", () => {
  const env = createMatchMediaStub({
    "(prefers-color-scheme: dark)": true,
    "(forced-colors: active)": true
  });
  const document = createStubDocument();
  const storage = createMemoryStorage();
  saveAppearanceSettings({ themePreference: "system" }, storage);
  const controller = new ThemeController({ env, document, storage });
  controller.start();

  assert.equal(controller.getThemePreference(), "system");
  assert.equal(document.documentElement.getAttribute("data-theme"), "high-contrast");
});

test("ThemeController falls back gracefully when storage is unavailable", () => {
  const env = createMatchMediaStub({
    "(prefers-color-scheme: dark)": true,
    "(forced-colors: active)": false
  });
  const document = createStubDocument();
  const storage = {
    getItem() {
      throw new Error("Storage disabled");
    },
    setItem() {
      throw new Error("Storage disabled");
    },
  };

  const controller = new ThemeController({ env, document, storage });
  controller.start();
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");

  // Mutations still apply even if persistence fails.
  controller.setThemePreference("dark");
  assert.equal(document.documentElement.getAttribute("data-theme"), "dark");
});
