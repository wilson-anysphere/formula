import test from "node:test";
import assert from "node:assert/strict";

import { ThemeController, resolveTheme } from "../src/theme/index.js";
import { createMemoryStorage } from "../src/settings/appearance/index.js";

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

test("ThemeController applies resolved theme + reduced motion and persists preference", () => {
  const env = createMatchMediaStub({
    "(prefers-color-scheme: dark)": true,
    "(prefers-reduced-motion: reduce)": false
  });
  const document = createStubDocument();
  const storage = createMemoryStorage();

  const controller = new ThemeController({ env, document, storage });
  controller.start();

  assert.equal(document.documentElement.getAttribute("data-theme"), "dark");
  assert.equal(document.documentElement.getAttribute("data-reduced-motion"), "false");

  controller.setThemePreference("light");
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");

  // System changes should not affect a user override.
  env.setMatches("(prefers-color-scheme: dark)", false);
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");

  // Persistence.
  const controller2 = new ThemeController({ env, document: createStubDocument(), storage });
  controller2.start();
  assert.equal(controller2.getThemePreference(), "light");
});

test("ThemeController updates theme when following system preferences", () => {
  const env = createMatchMediaStub({ "(prefers-color-scheme: dark)": false });
  const document = createStubDocument();
  const storage = createMemoryStorage();

  const controller = new ThemeController({ env, document, storage });
  controller.start();

  assert.equal(controller.getThemePreference(), "system");
  assert.equal(document.documentElement.getAttribute("data-theme"), "light");

  env.setMatches("(prefers-color-scheme: dark)", true);
  assert.equal(document.documentElement.getAttribute("data-theme"), "dark");
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
