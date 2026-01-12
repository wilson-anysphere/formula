import test from "node:test";
import assert from "node:assert/strict";

import { getDirection, setLocale, t, tWithVars } from "./index.js";

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

test("tWithVars interpolates placeholders", () => {
  setLocale("en-US");
  assert.equal(tWithVars("chat.errorWithMessage", { message: "Oops" }), "Error: Oops");
});

test("sheet position status bar string is localizable", () => {
  setLocale("en-US");
  assert.equal(tWithVars("statusBar.sheetPosition", { position: 1, total: 3 }), "Sheet 1 of 3");

  setLocale("de-DE");
  assert.equal(tWithVars("statusBar.sheetPosition", { position: 1, total: 3 }), "Blatt 1 von 3");

  setLocale("ar");
  assert.equal(tWithVars("statusBar.sheetPosition", { position: 1, total: 3 }), "ورقة 1 من 3");
});
