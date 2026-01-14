import assert from "node:assert/strict";
import test from "node:test";

import { ar } from "../src/i18n/locales/ar.js";
import { deDE } from "../src/i18n/locales/de-DE.js";
import { enUS } from "../src/i18n/locales/en-US.js";

function commandCategoryKeys(messages) {
  return Object.keys(messages)
    .filter((key) => key.startsWith("commandCategory."))
    .sort();
}

test("all locales define the same commandCategory.* keys as en-US", () => {
  const expected = commandCategoryKeys(enUS);
  const expectedSet = new Set(expected);

  const locales = {
    "en-US": enUS,
    "de-DE": deDE,
    ar,
  };

  for (const [locale, messages] of Object.entries(locales)) {
    const actual = commandCategoryKeys(messages);
    const actualSet = new Set(actual);

    const missing = expected.filter((key) => !Object.prototype.hasOwnProperty.call(messages, key));
    const extra = actual.filter((key) => !expectedSet.has(key));

    assert.deepStrictEqual(
      missing,
      [],
      `Locale ${locale} is missing command category keys: ${missing.length > 0 ? missing.join(", ") : "(none)"}`,
    );
    assert.deepStrictEqual(
      extra,
      [],
      `Locale ${locale} defines extra command category keys (add to en-US too): ${extra.length > 0 ? extra.join(", ") : "(none)"}`,
    );

    // Ensure values are strings (helps catch accidental non-string i18n edits).
    for (const key of expected) {
      assert.equal(typeof messages[key], "string", `Locale ${locale} command category ${key} must be a string`);
    }

    // Ensure ordering isn't masking duplicates (should be impossible, but keeps intent clear).
    assert.equal(actual.length, actualSet.size, `Locale ${locale} command category keys should be unique`);
  }
});

