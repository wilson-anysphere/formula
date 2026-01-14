<#
.SYNOPSIS
  Extract localized Excel error-literal spellings from a real Microsoft Excel install.

.DESCRIPTION
  This script launches Microsoft Excel via COM automation, writes formulas like `=#VALUE!`
  into a temporary workbook, and reads back the localized error literal shown by Excel for
  the active UI language.

  It then writes a TSV mapping:
    Canonical<TAB>Localized

  Canonical error literals are scraped from:
    crates/formula-engine/src/value/mod.rs (ErrorKind::as_code)

  By default the TSV is written to:
    crates/formula-engine/src/locale/data/upstream/errors/<locale>.tsv

  After extracting/updating an upstream TSV, regenerate the committed exports with:
    node scripts/generate-locale-error-tsvs.mjs

.PARAMETER Locale
  Locale tag to use for the output file name (e.g. "es-ES", "de-DE").

  If omitted, the script attempts to detect the Excel UI locale via:
    Application.LanguageSettings.LanguageID(msoLanguageIDUI)
  and converts it to a BCP-47 tag using .NET CultureInfo.

.PARAMETER OutPath
  Optional explicit output path. Overrides the default location derived from Locale.

.PARAMETER Visible
  Show Excel while running (useful for debugging).

.NOTES
  - Windows-only (requires Microsoft Excel desktop installed).
  - Excel versions can differ in which error literals they recognize (e.g. #SPILL!).
    If your Excel build does not support a newer error kind, this script may fail with
    an "unexpected error literal" message to avoid writing a misleading TSV.
#>

[CmdletBinding()]
param(
  [string]$Locale,
  [string]$OutPath,
  [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
  throw "extract-error-literals.ps1 is Windows-only (requires Excel desktop + COM automation)."
}

function Release-ComObject {
  param([object]$Object)
  if ($null -eq $Object) { return }
  try {
    if ([System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
      [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
  } catch {
    # Best-effort cleanup; ignore.
  }
}

function Set-RangeFormula {
  param(
    [Parameter(Mandatory = $true)]
    [object]$RangeObj,

    [Parameter(Mandatory = $true)]
    [string]$Formula
  )

  # Prefer `Formula2` when available (modern Excel) to avoid implicit-intersection quirks.
  try {
    $RangeObj.Formula2 = $Formula
  } catch {
    $RangeObj.Formula = $Formula
  }
}

function Extract-CanonicalErrorLiterals {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RustPath
  )

  if (-not (Test-Path -LiteralPath $RustPath)) {
    throw "Rust error-kind source not found: $RustPath"
  }

  $src = Get-Content -LiteralPath $RustPath -Raw -Encoding UTF8

  # Match lines like: `ErrorKind::Value => "#VALUE!",`
  $re = [regex]'ErrorKind::[A-Za-z0-9_]+\s*=>\s*"([^"]+)"'
  $matches = $re.Matches($src)
  if ($matches.Count -eq 0) {
    throw "Failed to extract any error literals from $RustPath (regex=$re)"
  }

  $out = New-Object System.Collections.Generic.List[string]
  $seen = @{}
  foreach ($m in $matches) {
    $code = [string]$m.Groups[1].Value
    if ($seen.ContainsKey($code)) {
      throw "Duplicate error literal found in $RustPath: $code"
    }
    $seen[$code] = $true
    $out.Add($code) | Out-Null
  }
  return ,$out.ToArray()
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot ".." "..")
$rustErrorKindPath = Join-Path $repoRoot "crates" "formula-engine" "src" "value" "mod.rs"
$canonicalCodes = Extract-CanonicalErrorLiterals -RustPath $rustErrorKindPath

$excel = $null
$workbook = $null
$sheet = $null
$cell = $null

$excelUiLocale = $null

try {
  try {
    $excel = New-Object -ComObject Excel.Application
  } catch {
    throw "Failed to create Excel COM object. Ensure Microsoft Excel is installed. Inner error: $($_.Exception.Message)"
  }

  $excel.Visible = [bool]$Visible
  $excel.DisplayAlerts = $false
  try { $excel.ScreenUpdating = $false } catch {}
  try { $excel.EnableEvents = $false } catch {}
  try { $excel.AskToUpdateLinks = $false } catch {}
  # msoAutomationSecurityForceDisable = 3 (disable macros)
  try { $excel.AutomationSecurity = 3 } catch {}

  try {
    # msoLanguageIDUI = 2
    $uiLcid = [int]$excel.LanguageSettings.LanguageID(2)
    $excelUiLocale = [System.Globalization.CultureInfo]::new($uiLcid).Name
  } catch {
    # Best effort; allow caller to pass -Locale explicitly.
    $excelUiLocale = $null
  }

  if (-not $Locale) {
    $Locale = $excelUiLocale
  }
  if (-not $Locale) {
    throw "Could not determine Excel UI locale. Pass -Locale (e.g. -Locale es-ES)."
  }
  if ($excelUiLocale -and -not ($excelUiLocale -ieq $Locale)) {
    Write-Warning "Excel UI locale '$excelUiLocale' does not match requested -Locale '$Locale'. Output reflects the active Excel UI language; ensure the correct Office language pack / display language is configured before extracting."
  }

  if (-not $OutPath) {
    $upstreamDir = Join-Path $repoRoot "crates" "formula-engine" "src" "locale" "data" "upstream" "errors"
    $OutPath = Join-Path $upstreamDir "$Locale.tsv"
  }

  $excelVersion = $null
  $excelBuild = $null
  try { $excelVersion = [string]$excel.Version } catch {}
  try { $excelBuild = [string]$excel.Build } catch {}

  Write-Host "Excel: version=$excelVersion build=$excelBuild uiLocale=$excelUiLocale"
  Write-Host "Extracting $($canonicalCodes.Count) error literals -> $OutPath"

  $workbook = $excel.Workbooks.Add()
  $sheet = $workbook.Worksheets.Item(1)
  $sheet.Name = "ErrorLiterals"
  # Ensure `.Text` has enough room if we need to fall back to it for longer literals.
  try { $sheet.Columns.Item(1).ColumnWidth = 60 } catch {}

  $cell = $sheet.Range("A1")

  $rows = New-Object System.Collections.Generic.List[object]
  foreach ($code in $canonicalCodes) {
    $formula = "=" + $code
    Set-RangeFormula -RangeObj $cell -Formula $formula
    # Error constants should not require recalculation, but some builds might defer updates.
    try { $excel.Calculate() } catch {}

    $formulaLocal = $null
    try { $formulaLocal = [string]$cell.FormulaLocal } catch { $formulaLocal = $null }
    $displayText = $null
    try { $displayText = [string]$cell.Text } catch { $displayText = $null }

    $candidate = $formulaLocal
    if ($candidate -is [string]) {
      $candidate = $candidate.Trim()
      # Defensive: Excel sometimes serializes formulas with extra leading markers
      # like `=+...` or `=@...`. Strip these before inspecting prefixes.
      while ($candidate.Length -gt 0) {
        $ch = $candidate.Substring(0, 1)
        if ($ch -eq "=" -or $ch -eq "@" -or $ch -eq "+") {
          $candidate = $candidate.Substring(1)
          continue
        }
        break
      }
    }
    if ($candidate -is [string]) {
      $candidate = $candidate.Trim()
    }

    $localized = $null
    if ($candidate -is [string] -and $candidate.StartsWith("#")) {
      $localized = $candidate
      # If FormulaLocal didn't translate the literal but the displayed value did, trust `.Text`.
      if (
        $candidate -ceq $code -and
        $displayText -is [string] -and
        $displayText.StartsWith("#") -and
        -not ($displayText -ceq $code)
      ) {
        $localized = $displayText
      }
    } elseif ($displayText -is [string] -and $displayText.StartsWith("#")) {
      $localized = $displayText
    } else {
      throw "Failed to extract error literal for $code (FormulaLocal=$formulaLocal, Text=$displayText)"
    }

    # Guardrail: for most errors, the trailing punctuation should be stable across locales.
    $last = $code.Substring($code.Length - 1, 1)
    if (($last -eq "!") -or ($last -eq "?")) {
      if (-not $localized.EndsWith($last)) {
        throw "Excel returned unexpected error literal for $code: $localized (expected to end with '$last'). This may indicate your Excel build does not recognize $code."
      }
    }

    $rows.Add([ordered]@{ canonical = $code; localized = $localized }) | Out-Null
  }

  $outLines = New-Object System.Collections.Generic.List[string]
  $outLines.Add("# Canonical`tLocalized") | Out-Null
  $outLines.Add("# Source: Extracted from Microsoft Excel via COM automation (tools/excel-oracle/extract-error-literals.ps1).") | Out-Null
  if ($excelUiLocale) {
    $outLines.Add("# Excel UI locale: $excelUiLocale") | Out-Null
  }
  $outLines.Add("#") | Out-Null
  foreach ($r in $rows) {
    $outLines.Add("$($r.canonical)`t$($r.localized)") | Out-Null
  }

  $outDir = Split-Path -Parent $OutPath
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  }

  $fullOutPath = [System.IO.Path]::GetFullPath($OutPath)
  [System.IO.File]::WriteAllText(
    $fullOutPath,
    ($outLines -join "`n") + "`n",
    [System.Text.UTF8Encoding]::new($false)
  )

  Write-Host "Wrote: $fullOutPath"
} finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($false) } catch {}
  }
  if ($null -ne $excel) {
    try { $excel.Quit() } catch {}
  }

  Release-ComObject $cell
  Release-ComObject $sheet
  Release-ComObject $workbook
  Release-ComObject $excel

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
