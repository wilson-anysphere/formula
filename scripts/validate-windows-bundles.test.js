import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "validate-windows-bundles.ps1");
const text = readFileSync(scriptPath, "utf8");

test("validate-windows-bundles.ps1 inspects MSI tables via Windows Installer COM", () => {
  assert.match(
    text,
    /New-Object\s+-ComObject\s+WindowsInstaller\.Installer/,
    "Expected validator to instantiate WindowsInstaller.Installer COM object for MSI inspection.",
  );
  assert.match(
    text,
    /OpenDatabase\(/,
    "Expected validator to open MSI databases (OpenDatabase) for table queries.",
  );
  assert.match(
    text,
    /FROM\s+`Extension`/,
    "Expected validator to query the MSI Extension table (file association evidence).",
  );
});

test("validate-windows-bundles.ps1 validates formula:// URL protocol handler via Registry table", () => {
  assert.match(
    text,
    /Assert-MsiRegistersUrlProtocol/,
    "Expected validator to include an MSI URL protocol handler assertion.",
  );
  assert.match(
    text,
    /FROM\s+`Registry`/,
    "Expected validator to query the MSI Registry table for protocol handler entries.",
  );
  assert.match(
    text,
    /URL Protocol/,
    "Expected validator to require the 'URL Protocol' registry value for URL protocol handler registration.",
  );
});

test("validate-windows-bundles.ps1 performs best-effort NSIS marker scanning", () => {
  assert.match(
    text,
    /Find-BinaryMarkerInFile/,
    "Expected validator to include streaming substring search for EXE markers (best-effort NSIS validation).",
  );
  assert.match(
    text,
    /x-scheme-handler\//,
    "Expected validator to scan for x-scheme-handler/<scheme> markers in NSIS installers.",
  );
  assert.match(
    text,
    /\.xlsx/i,
    "Expected validator to scan for .xlsx marker strings in NSIS installers.",
  );
});

test("validate-windows-bundles.ps1 accepts common Excel ProgId evidence for .xlsx associations", () => {
  assert.match(
    text,
    /Excel\.Sheet\.12/,
    "Expected validator to accept Excel.Sheet.12-style ProgId evidence for .xlsx associations (Registry fallback).",
  );
});

test("validate-windows-bundles.ps1 prefers validating the stable formula:// scheme when multiple schemes exist", () => {
  assert.match(
    text,
    /-contains\s+\"formula\"/,
    "Expected validator to explicitly prefer validating the 'formula' URL scheme when it exists in the configured scheme list.",
  );
});

test("validate-windows-bundles.ps1 validates MSI UpgradeCode against tauri.conf.json (WiX upgrades/downgrades)", () => {
  assert.match(
    text,
    /Get-ExpectedWixUpgradeCode/,
    "Expected validator to read bundle.windows.wix.upgradeCode from tauri.conf.json.",
  );
  assert.match(
    text,
    /PropertyName\s+\"UpgradeCode\"/,
    "Expected validator to query the MSI Property table for UpgradeCode.",
  );
  assert.match(
    text,
    /Normalize-Guid/,
    "Expected validator to normalize GUID formatting when comparing UpgradeCode values.",
  );
});
