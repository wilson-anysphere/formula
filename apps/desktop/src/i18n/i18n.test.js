import test from "node:test";
import assert from "node:assert/strict";

import { getDirection, setLocale, t, tWithVars } from "./index.js";

test("switching locale changes translations", () => {
  setLocale("en-US");
  assert.equal(t("menu.file"), "File");

  setLocale("de-DE");
  assert.equal(t("menu.file"), "Datei");
});

test("number format command strings are localizable", () => {
  setLocale("en-US");
  assert.equal(t("command.format.numberFormat.general"), "General");

  setLocale("de-DE");
  assert.equal(t("command.format.numberFormat.general"), "Allgemein");

  setLocale("ar");
  assert.equal(t("command.format.numberFormat.general"), "عام");
});

test("theme command strings are localizable", () => {
  setLocale("en-US");
  assert.equal(t("command.view.theme.light"), "Theme: Light");

  setLocale("de-DE");
  assert.equal(t("command.view.theme.light"), "Design: Hell");

  setLocale("ar");
  assert.equal(t("command.view.theme.light"), "السمة: فاتح");
});

test("ribbon label strings are localizable", () => {
  setLocale("en-US");
  assert.equal(t("ribbon.label.mixed"), "Mixed");
  assert.equal(t("ribbon.label.custom"), "Custom");
  assert.equal(t("ribbon.label.comma"), "Comma");

  setLocale("de-DE");
  assert.equal(t("ribbon.label.mixed"), "Gemischt");
  assert.equal(t("ribbon.label.custom"), "Benutzerdefiniert");
  assert.equal(t("ribbon.label.comma"), "Komma");

  setLocale("ar");
  assert.equal(t("ribbon.label.mixed"), "مختلط");
  assert.equal(t("ribbon.label.custom"), "مخصص");
  assert.equal(t("ribbon.label.comma"), "فاصلة");
});

test("number format quick-pick strings are localizable", () => {
  setLocale("en-US");
  assert.equal(t("quickPick.numberFormat.placeholder"), "Number format");
  assert.equal(t("prompt.customNumberFormat.code"), "Custom number format code");
  assert.equal(t("command.home.number.moreFormats.custom"), "Custom number format…");

  setLocale("de-DE");
  assert.equal(t("quickPick.numberFormat.placeholder"), "Zahlenformat");
  assert.equal(t("prompt.customNumberFormat.code"), "Benutzerdefinierter Zahlenformatcode");
  assert.equal(t("command.home.number.moreFormats.custom"), "Benutzerdefiniertes Zahlenformat…");

  setLocale("ar");
  assert.equal(t("quickPick.numberFormat.placeholder"), "تنسيق الأرقام");
  assert.equal(t("prompt.customNumberFormat.code"), "رمز تنسيق الأرقام المخصص");
  assert.equal(t("command.home.number.moreFormats.custom"), "تنسيق أرقام مخصص…");
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

test("Version History Compare strings are localizable", () => {
  setLocale("en-US");
  assert.equal(t("versionHistory.compare.title"), "Compare");
  assert.equal(tWithVars("versionHistory.compare.badge.added", { count: 2 }), "Added: 2");

  setLocale("de-DE");
  assert.equal(t("versionHistory.compare.title"), "Vergleichen");
  assert.equal(tWithVars("versionHistory.compare.badge.added", { count: 2 }), "Hinzugefügt: 2");

  setLocale("ar");
  assert.equal(t("versionHistory.compare.title"), "مقارنة");
  assert.equal(tWithVars("versionHistory.compare.badge.added", { count: 2 }), "تمت الإضافة: 2");
});
