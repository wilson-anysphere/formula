import test from "node:test";
import assert from "node:assert/strict";

import { getDirection, setLocale, t } from "./index.js";

test("switching locale changes translations", () => {
  setLocale("en-US");
  assert.equal(t("menu.file"), "File");

  setLocale("de-DE");
  assert.equal(t("menu.file"), "Datei");
});

test("rtl locale exposes rtl direction hook", () => {
  setLocale("ar");
  assert.equal(getDirection(), "rtl");
});

