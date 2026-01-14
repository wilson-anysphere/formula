import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { renderAppShell } from "../test/themeScreensSnapshotUtils.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const snapshotDir = path.join(__dirname, "..", "test", "__snapshots__");

const themes = [
  { theme: "light", file: "app-shell.light.html" },
  { theme: "dark", file: "app-shell.dark.html" },
  { theme: "high-contrast", file: "app-shell.high-contrast.html" },
];

for (const { theme, file } of themes) {
  const html = renderAppShell(theme);
  const outPath = path.join(snapshotDir, file);
  fs.writeFileSync(outPath, html);
  console.log(`Wrote ${path.relative(process.cwd(), outPath)} (${html.length.toLocaleString()} chars)`);
}

