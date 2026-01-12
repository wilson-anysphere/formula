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

test("Tauri main capability allows emitting coi-check-result (used by pnpm check:coi)", () => {
  const capPath = path.join(__dirname, "..", "src-tauri", "capabilities", "main.json");
  const cap = JSON.parse(fs.readFileSync(capPath, "utf8"));

  const perms = cap?.permissions;
  assert.ok(Array.isArray(perms), "expected capabilities/main.json to have a permissions array");

  // Tauri v2.9 core permissions use the `core:` prefix (see `cargo tauri permission ls`).
  const emitPerm = perms.find(
    (p) => p && typeof p === "object" && p.identifier === "core:event:allow-emit",
  );
  assert.ok(
    emitPerm,
    "expected capabilities/main.json to include a core:event:allow-emit permission object",
  );

  const allowed = Array.isArray(emitPerm.allow) ? emitPerm.allow : [];
  const allowedEvents = allowed.map((entry) => entry?.event).filter(Boolean);
  assert.ok(
    allowedEvents.includes("coi-check-result"),
    "expected capabilities/main.json to allow emitting event 'coi-check-result' (required for the packaged cross-origin isolation smoke check)",
  );
});
