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
  [Alias("LocaleId")]
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

function Parse-FormulaLocalErrorLiteral {
  param([string]$FormulaLocal)

  if (-not $FormulaLocal) { return $null }

  $candidate = $FormulaLocal.Trim()
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
  $candidate = $candidate.Trim()

  if ($candidate.StartsWith("#")) {
    return $candidate
  }
  return $null
}

function Expected-SentinelErrorTranslations {
  param([Parameter(Mandatory = $true)][string]$LocaleId)

  switch ($LocaleId) {
    "de-DE" {
      return @{
        "#VALUE!" = "#WERT!"
        "#REF!" = "#BEZUG!"
        "#SPILL!" = "#ÜBERLAUF!"
        "#GETTING_DATA" = "#DATEN_ABRUFEN"
      }
    }
    "fr-FR" {
      return @{
        "#VALUE!" = "#VALEUR!"
        "#NAME?" = "#NOM?"
        "#GETTING_DATA" = "#OBTENTION_DONNEES"
      }
    }
    "es-ES" {
      return @{
        "#VALUE!" = "#¡VALOR!"
        "#NAME?" = "#¿NOMBRE?"
        "#GETTING_DATA" = "#OBTENIENDO_DATOS"
      }
    }
    default { return $null }
  }
}

function Warn-IfExcelErrorLocaleSeemsMisconfigured {
  param(
    [Parameter(Mandatory = $true)][object]$ExcelObj,
    [Parameter(Mandatory = $true)][object]$RangeObj,
    [Parameter(Mandatory = $true)][string]$LocaleId
  )

  $expected = Expected-SentinelErrorTranslations -LocaleId $LocaleId
  if ($null -eq $expected) { return }

  foreach ($kv in $expected.GetEnumerator()) {
    $canon = [string]$kv.Key
    $want = [string]$kv.Value
    try {
      $RangeObj.Clear()
      Set-RangeFormula -RangeObj $RangeObj -Formula ("=" + $canon)
      try { $ExcelObj.Calculate() } catch {}

      $text = $null
      try { $text = [string]$RangeObj.Text } catch { $text = $null }
      $formulaLocal = $null
      try { $formulaLocal = [string]$RangeObj.FormulaLocal } catch { $formulaLocal = $null }

      $got = $null
      if ($text -is [string] -and $text.StartsWith("#")) {
        $got = $text
      } else {
        $got = Parse-FormulaLocalErrorLiteral -FormulaLocal $formulaLocal
      }

      if (-not $got) {
        Write-Warning "Could not sanity-check Excel error literal translation for $canon (Text=$text, FormulaLocal=$formulaLocal)."
        return
      }
      if (-not ($got -ieq $want)) {
        Write-Warning "Excel locale may be misconfigured: expected $canon -> $want for locale '$LocaleId', got '$got'. (This script reflects the active Excel UI language.)"
        return
      }
    } catch {
      Write-Warning "Failed sanity-checking Excel error literal translation for $canon: $($_.Exception.Message)"
      return
    }
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

$outPathIsTsvFile = $false
if ($OutPath) {
  $outPathIsTsvFile = $OutPath.ToLowerInvariant().EndsWith(".tsv")
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
  } elseif (-not $outPathIsTsvFile) {
    # Match `extract-function-translations.ps1` behavior: allow callers to pass a directory.
    $OutPath = Join-Path $OutPath "$Locale.tsv"
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

  Warn-IfExcelErrorLocaleSeemsMisconfigured -ExcelObj $excel -RangeObj $cell -LocaleId $Locale

  # Guardrail: if we ever get the same displayed error for multiple canonical codes,
  # it likely means this Excel build doesn't recognize one of the newer error kinds
  # and substituted a different error (often #NAME?). Fail rather than emitting an
  # ambiguous/incorrect mapping.
  $seenLocalized = @{}

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

    $candidate = Parse-FormulaLocalErrorLiteral -FormulaLocal $formulaLocal

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

    # Guardrail: if Excel substituted a *different* canonical error literal (e.g. returning `#NAME?`
    # when asked for `#GETTING_DATA`), fail rather than writing a misleading mapping. This can
    # happen when a given Excel build does not recognize a newer error kind.
    foreach ($other in $canonicalCodes) {
      if (($other -ieq $localized) -and -not ($other -ieq $code)) {
        throw "Excel returned canonical error literal $localized when extracting $code (expected the same error kind). This may indicate your Excel build does not recognize $code."
      }
    }

    $folded = $localized.ToUpperInvariant()
    if ($seenLocalized.ContainsKey($folded) -and -not ($seenLocalized[$folded] -ieq $code)) {
      $prev = [string]$seenLocalized[$folded]
      throw "Excel returned the same localized error literal for multiple canonical codes: $prev and $code both mapped to $localized. This may indicate your Excel build does not recognize one of the error kinds."
    }
    $seenLocalized[$folded] = $code

    # Guardrail: for most errors, the trailing punctuation should be stable across locales.
    $last = $code.Substring($code.Length - 1, 1)
    if (($last -eq "!") -or ($last -eq "?")) {
      if (-not $localized.EndsWith($last)) {
        throw "Excel returned unexpected error literal for $code: $localized (expected to end with '$last'). This may indicate your Excel build does not recognize $code."
      }
    }

    Write-Verbose ("{0} -> {1} (FormulaLocal={2}, Text={3})" -f $code, $localized, $formulaLocal, $displayText)

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
