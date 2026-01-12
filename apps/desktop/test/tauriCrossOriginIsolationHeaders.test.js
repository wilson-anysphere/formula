import fs from "node:fs";
import path from "node:path";
import test from "node:test";
import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test("packaged Tauri config enables cross-origin isolation headers (COOP/COEP)", () => {
  const configPath = path.join(__dirname, "..", "src-tauri", "tauri.conf.json");
  const config = JSON.parse(fs.readFileSync(configPath, "utf8"));

  const headers = config?.app?.security?.headers;
  assert.ok(headers && typeof headers === "object", "expected app.security.headers to be present in tauri.conf.json");

  assert.equal(
    headers["Cross-Origin-Opener-Policy"],
    "same-origin",
    "Packaged Tauri builds must set COOP to enable globalThis.crossOriginIsolated (SharedArrayBuffer / Pyodide worker backend).",
  );
  assert.equal(
    headers["Cross-Origin-Embedder-Policy"],
    "require-corp",
    "Packaged Tauri builds must set COEP=require-corp to enable globalThis.crossOriginIsolated (SharedArrayBuffer / Pyodide worker backend).",
  );
});

