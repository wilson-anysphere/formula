import test from "node:test";

import { expectSnapshot } from "./snapshot.js";
import { renderAppShell } from "./themeScreensSnapshotUtils.js";

test("app shell snapshot (light)", () => {
  expectSnapshot("app-shell.light.html", renderAppShell("light"));
});

test("app shell snapshot (dark)", () => {
  expectSnapshot("app-shell.dark.html", renderAppShell("dark"));
});

test("app shell snapshot (high contrast)", () => {
  expectSnapshot("app-shell.high-contrast.html", renderAppShell("high-contrast"));
});
