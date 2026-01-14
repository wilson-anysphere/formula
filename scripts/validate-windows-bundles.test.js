import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripPowerShellComments } from "../apps/desktop/test/sourceTextUtils.js";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const scriptPath = path.join(repoRoot, "scripts", "validate-windows-bundles.ps1");
const text = stripPowerShellComments(readFileSync(scriptPath, "utf8"));

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
    /\$dotExt\s*=\s*"\."\s*\+/,
    "Expected validator to build dotted file-extension markers (e.g. .xlsx) for NSIS installer scanning.",
  );
  assert.match(
    text,
    /Extensions\s*=\s*@\(\s*"xlsx"\s*\)/,
    "Expected validator to include a stable representative extension (xlsx) for best-effort NSIS marker scanning.",
  );
});

test("validate-windows-bundles.ps1 uses token-bounded x-scheme-handler markers (avoid prefix false positives)", () => {
  assert.match(
    text,
    /x-scheme-handler\/\$scheme;/,
    "Expected validator to search for x-scheme-handler/<scheme>; (semicolon-delimited) to avoid matching x-scheme-handler/<scheme>-extra.",
  );
  assert.doesNotMatch(
    text,
    /\"x-scheme-handler\/\$scheme\"\s*,/,
    "Expected validator to avoid bare x-scheme-handler/$scheme markers without a token boundary.",
  );
});

test("validate-windows-bundles.ps1 best-effort URL protocol scan requires scheme-specific markers (not just 'URL Protocol')", () => {
  const match = text.match(/function\s+Find-ExeUrlProtocolMarker[\s\S]*?\n\s*}\s*\n/s);
  assert.ok(match, "Expected Find-ExeUrlProtocolMarker function to exist");
  const body = match[0];
  // This scan is heuristic, but it should be anchored to scheme-specific registry paths to avoid
  // prefix false positives (e.g. avoid treating formula-extra as satisfying formula).
  assert.ok(
    body.includes("shell\\open\\command"),
    "Expected Find-ExeUrlProtocolMarker to scan for <scheme>\\\\shell\\\\open\\\\command markers.",
  );
  assert.doesNotMatch(
    body,
    /\"URL Protocol\"/,
    "Expected Find-ExeUrlProtocolMarker to avoid using generic 'URL Protocol' markers without the scheme name.",
  );
});

test("validate-windows-bundles.ps1 MSI URL protocol fallback scan is scheme-specific (no prefix matches)", () => {
  const match = text.match(/\$schemeNeedles\s*=\s*@\([\s\S]*?\)\s*\n\s*\$schemeMarker\s*=\s*Find-BinaryMarkerInFile/s);
  assert.ok(match, "Expected MSI URL protocol fallback to define $schemeNeedles for marker scanning.");
  const block = match[0];
  assert.ok(
    block.includes("shell\\open\\command"),
    "Expected MSI URL protocol fallback needles to include shell\\\\open\\\\command markers.",
  );
  assert.doesNotMatch(
    block,
    /\"URL Protocol\"/,
    "Expected MSI URL protocol fallback needles to avoid generic 'URL Protocol' markers without the scheme name.",
  );
});

test("validate-windows-bundles.ps1 MSI file association fallback avoids bare extension markers", () => {
  const match = text.match(/\$needles\s*=\s*@\([\s\S]*?\)\s*\n\s*if\s*\(\s*\$requiredExtensionNoDot\s+-ieq\s+\"xlsx\"/s);
  assert.ok(match, "Expected MSI file association fallback to define $needles marker list.");
  const block = match[0];
  // The fallback should prefer registry-path context rather than just scanning for ".xlsx" or "xlsx"
  // anywhere in the MSI binary.
  assert.doesNotMatch(block, /\n\s*\$dotExt\s*,?\s*\n/, "Expected file association fallback to avoid bare $dotExt entries.");
  assert.doesNotMatch(
    block,
    /\n\s*\$requiredExtensionNoDot\s*,?\s*\n/,
    "Expected file association fallback to avoid bare $requiredExtensionNoDot entries.",
  );
});

test("validate-windows-bundles.ps1 validates all configured file association extensions (not just .xlsx) via MSI", () => {
  assert.match(
    text,
    /expectedExtensions\s*=\s*@\(/,
    "Expected validator to derive an extension list from tauri.conf.json bundle.fileAssociations.",
  );
  assert.match(
    text,
    /foreach\s*\(\s*\$ext\s+in\s+\$expectedExtensions\s*\)/,
    "Expected validator to loop over all configured extensions when validating MSI file association metadata.",
  );
  assert.match(
    text,
    /Assert-MsiDeclaresFileAssociation\s+-Msi\s+\$msi\s+-ExtensionNoDot\s+\$ext/,
    "Expected validator to validate each extension via the MSI Extension table.",
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

test("validate-windows-bundles.ps1 validates all configured URL protocol schemes via MSI", () => {
  assert.match(
    text,
    /foreach\s*\(\s*\$scheme\s+in\s+\$expectedSchemes\s*\)/i,
    "Expected validator to loop over all configured deep-link schemes when validating MSI URL protocol registration.",
  );
  assert.match(
    text,
    /Assert-MsiRegistersUrlProtocol\s+-Msi\s+\$msi\s+-Scheme\s+\$scheme/i,
    "Expected validator to validate each URL scheme via Assert-MsiRegistersUrlProtocol.",
  );
});

test("validate-windows-bundles.ps1 rejects invalid deep-link schemes in tauri.conf.json (e.g. formula://evil)", () => {
  assert.match(
    text,
    /Invalid deep-link scheme configured in tauri\.conf\.json/i,
    "Expected validator to throw a clear error when plugins.deep-link.desktop.schemes contains invalid values with ':' or '/'.",
  );
  // Ensure the check is present (not just the error string).
  assert.ok(
    text.includes("$v -match '[:/]'"),
    "Expected validator to check for invalid characters in normalized schemes (contains $v -match '[:/]').",
  );
});

test("validate-windows-bundles.ps1 asserts LICENSE/NOTICE are included in MSI installers", () => {
  assert.match(
    text,
    /Assert-MsiContainsComplianceArtifacts/,
    "Expected validator to include an MSI compliance artifacts assertion (LICENSE/NOTICE).",
  );
  assert.match(text, /Compliance artifact check \(MSI\)/);
  assert.match(text, /\$required\s*=\s*@\(\"LICENSE\",\s*\"NOTICE\"\)/);
});

test("validate-windows-bundles.ps1 asserts LICENSE/NOTICE are included in NSIS (.exe) installer payloads", () => {
  assert.match(
    text,
    /Assert-ExeContainsComplianceArtifacts/,
    "Expected validator to include an EXE compliance artifacts assertion (LICENSE/NOTICE).",
  );
  // Payload inspection relies on 7-Zip extraction when available.
  assert.match(text, /7-Zip/i);
  assert.match(text, /foreach\s*\(\$req\s+in\s+@\(\"LICENSE\",\s*\"NOTICE\"\)\)/);
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

test("validate-windows-bundles.ps1 avoids unbounded Get-ChildItem -Recurse scans over target roots (perf guardrail)", () => {
  // Ensure we don't regress back to a full `Get-ChildItem -Recurse` over `target/` when
  // bundle discovery fails. Cargo target directories can be enormous once builds have run.
  assert.doesNotMatch(
    text,
    /Get-ChildItem\s+-LiteralPath\s+\\$TargetRoot\s+-Recurse\s+-Directory/,
    "Expected bundle discovery to avoid unbounded `Get-ChildItem -LiteralPath $TargetRoot -Recurse -Directory` scans.",
  );
});

test("validate-windows-bundles.ps1 validates MSI ProductName against tauri.conf.json productName", () => {
  assert.match(
    text,
    /Get-ExpectedProductName/,
    "Expected validator to read productName from tauri.conf.json.",
  );
  assert.match(
    text,
    /PropertyName\s+\"ProductName\"/,
    "Expected validator to query the MSI Property table for ProductName.",
  );
});

test("validate-windows-bundles.ps1 supports overriding tauri.conf.json path via FORMULA_TAURI_CONF_PATH", () => {
  assert.match(
    text,
    /FORMULA_TAURI_CONF_PATH/,
    "Expected validator to support FORMULA_TAURI_CONF_PATH override (consistent with other desktop validation scripts).",
  );
  assert.match(
    text,
    /Get-TauriConfPath/,
    "Expected validator to centralize tauri.conf.json path resolution via a helper (Get-TauriConfPath).",
  );
});
