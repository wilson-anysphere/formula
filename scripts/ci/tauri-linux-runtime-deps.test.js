import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const tauriConfPath = path.join(repoRoot, "apps", "desktop", "src-tauri", "tauri.conf.json");

/** @type {any} */
const tauriConf = JSON.parse(readFileSync(tauriConfPath, "utf8"));

/**
 * @param {unknown} deps
 * @returns {string[]}
 */
function normalizeDeps(deps) {
  if (!Array.isArray(deps)) return [];
  return deps.map((v) => String(v));
}

/**
 * @param {string[]} deps
 * @param {RegExp} re
 */
function assertAnyMatch(deps, re) {
  assert.ok(
    deps.some((d) => re.test(d)),
    `Expected at least one dependency matching ${re}.\nFound:\n- ${deps.join("\n- ")}`,
  );
}

test("tauri.conf.json declares Linux .deb runtime dependencies (WebKitGTK/GTK/AppIndicator/librsvg/OpenSSL)", () => {
  const deps = normalizeDeps(tauriConf?.bundle?.linux?.deb?.depends);
  assert.ok(deps.length > 0, "bundle.linux.deb.depends is missing/empty in tauri.conf.json");

  // Require WebKitGTK 4.1 explicitly (avoid drifting to 4.0).
  assertAnyMatch(deps, /libwebkit2gtk-4\.1/i);
  assertAnyMatch(deps, /libgtk-3/i);
  assertAnyMatch(deps, /appindicator/i);
  assertAnyMatch(deps, /librsvg2/i);
  assertAnyMatch(deps, /libssl/i);
});

test("tauri.conf.json declares Linux .rpm runtime dependencies (WebKitGTK/GTK/AppIndicator/librsvg/OpenSSL)", () => {
  const deps = normalizeDeps(tauriConf?.bundle?.linux?.rpm?.depends);
  assert.ok(deps.length > 0, "bundle.linux.rpm.depends is missing/empty in tauri.conf.json");

  // We use RPM rich dependencies (`(a or b)`) to cover Fedora/RHEL + openSUSE naming differences.
  // Reject common copy/paste mistakes from Debian-style dependency syntax.
  assert.ok(
    deps.every((d) => !d.includes("|")),
    `bundle.linux.rpm.depends must not use Debian-style '|' alternation.\nFound:\n- ${deps.join("\n- ")}`,
  );
  assert.ok(
    deps.every((d) => !/\bt64\b/i.test(d)),
    `bundle.linux.rpm.depends must not reference Ubuntu/Debian t64 package variants.\nFound:\n- ${deps.join("\n- ")}`,
  );

  assertAnyMatch(deps, /webkit2gtk4\.1/i);
  assertAnyMatch(deps, /libwebkit2gtk-4_1/i);
  assertAnyMatch(deps, /\bgtk3\b/i);
  assertAnyMatch(deps, /libgtk-3-0/i);
  assertAnyMatch(deps, /appindicator/i);
  assertAnyMatch(deps, /librsvg/i);
  assertAnyMatch(deps, /openssl/i);
});
