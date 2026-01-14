<#
.SYNOPSIS
  Probe how Excel formats numeric literals in `Range.FormulaLocal`.

.DESCRIPTION
  A regression introduced locale thousands-grouping insertion when localizing formulas in
  `crates/formula-engine/src/locale/translate.rs`. Excel's behavior here is subtle:

  - Excel may ACCEPT thousands separators in localized formula input (e.g. `1.234,56` in `de-DE`)
  - But `FormulaLocal` may NORMALIZE numeric literals when serializing the formula back (often
    omitting grouping separators entirely)

  This script automates a small set of numeric-heavy formulas via Excel COM automation and records:

  - the formula set via `Range.Formula` (canonical/en-US form)
  - the observed `Range.FormulaLocal`
  - the result of attempting to set `Range.FormulaLocal` with explicit grouping separators and
    reading back `Range.FormulaLocal` to detect normalization

  It runs the probe under several separator configurations that correspond to common locales:
  `en-US`, `de-DE`, `fr-FR`, `es-ES`.

  NOTE:
  - Excel COM automation is Windows-only and requires Excel desktop installed.
  - This script does NOT (and generally cannot) switch Excel's UI language at runtime. Function
    names in `FormulaLocal` reflect the installed Excel UI language, but numeric punctuation is
    controlled by `Application.UseSystemSeparators` + `DecimalSeparator`/`ThousandsSeparator`.

.PARAMETER OutPath
  Output path for the generated JSON.

  - If OutPath ends with ".json", it is treated as a file path.
  - Otherwise, it is treated as a directory, and "formula-local-number-formatting.json" is written inside it.

.PARAMETER Visible
  Make Excel visible while running (useful for debugging).

.EXAMPLE
  # From repo root on Windows (requires Excel desktop installed)
  powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-formula-local-number-formatting.ps1 `
    -OutPath out/
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
  throw "extract-formula-local-number-formatting.ps1 is Windows-only (requires Excel desktop + COM automation)."
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

function Resolve-FullPath {
  param([Parameter(Mandatory = $true)][string]$Path)
  if ([System.IO.Path]::IsPathRooted($Path)) {
    return [System.IO.Path]::GetFullPath($Path)
  }
  return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function Ensure-OutputFile {
  param([Parameter(Mandatory = $true)][string]$OutPath)

  $full = Resolve-FullPath -Path $OutPath
  if ($full.ToLowerInvariant().EndsWith(".json")) {
    $dir = Split-Path -Parent $full
    if ($dir -and !(Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    return $full
  }

  if (!(Test-Path $full)) { [void](New-Item -ItemType Directory -Path $full -Force) }
  return (Join-Path $full "formula-local-number-formatting.json")
}

function Set-RangeFormula {
  param(
    [Parameter(Mandatory = $true)][object]$RangeObj,
    [Parameter(Mandatory = $true)][string]$Formula
  )

  # Prefer `Formula2` when available (modern Excel) to avoid implicit-intersection quirks.
  try {
    $RangeObj.Formula2 = $Formula
  } catch {
    $RangeObj.Formula = $Formula
  }
}

function Insert-ThousandsSeparators {
  param(
    [Parameter(Mandatory = $true)][string]$Digits,
    [Parameter(Mandatory = $true)][string]$Sep
  )

  if ($Digits.Length -le 3) { return $Digits }

  $out = ""
  $i = $Digits.Length
  while ($i -gt 3) {
    $out = $Sep + $Digits.Substring($i - 3, 3) + $out
    $i -= 3
  }
  $Digits.Substring(0, $i) + $out
}

function Add-GroupingToNumberLiteral {
  param(
    [Parameter(Mandatory = $true)][string]$Raw,
    [Parameter(Mandatory = $true)][string]$DecimalSeparator,
    [Parameter(Mandatory = $true)][string]$ThousandsSeparator
  )

  $mantissa = $Raw
  $exponent = ""

  $eIdx = $Raw.IndexOf("E")
  if ($eIdx -lt 0) { $eIdx = $Raw.IndexOf("e") }
  if ($eIdx -ge 0) {
    $mantissa = $Raw.Substring(0, $eIdx)
    $exponent = $Raw.Substring($eIdx)
  }

  $sign = ""
  if ($mantissa.StartsWith("+") -or $mantissa.StartsWith("-")) {
    $sign = $mantissa.Substring(0, 1)
    $mantissa = $mantissa.Substring(1)
  }

  $intPart = $mantissa
  $fracPart = ""
  $decIdx = $mantissa.IndexOf($DecimalSeparator)
  if ($decIdx -ge 0) {
    $intPart = $mantissa.Substring(0, $decIdx)
    $fracPart = $mantissa.Substring($decIdx)
  }

  if ($intPart.Length -le 3) { return $Raw }
  if ($intPart -notmatch '^\d+$') { return $Raw }

  $grouped = Insert-ThousandsSeparators -Digits $intPart -Sep $ThousandsSeparator
  $sign + $grouped + $fracPart + $exponent
}

function Add-GroupingToFormulaLocal {
  param(
    [Parameter(Mandatory = $true)][string]$FormulaLocal,
    [Parameter(Mandatory = $true)][string]$DecimalSeparator,
    [Parameter(Mandatory = $true)][string]$ThousandsSeparator
  )

  $decEsc = [regex]::Escape($DecimalSeparator)
  $pattern = "(?<![A-Za-z0-9_])([+-]?\d+(?:$decEsc\d+)?(?:[Ee][+-]?\d+)?)(?![A-Za-z0-9_])"

  [regex]::Replace(
    $FormulaLocal,
    $pattern,
    {
      param($m)
      Add-GroupingToNumberLiteral -Raw $m.Groups[1].Value -DecimalSeparator $DecimalSeparator -ThousandsSeparator $ThousandsSeparator
    }
  )
}

$outFile = Ensure-OutputFile -OutPath $OutPath

$excel = $null
$workbook = $null
$sheet = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = [bool]$Visible
  $excel.DisplayAlerts = $false

  $origUseSystemSeparators = $null
  $origDecimalSeparator = $null
  $origThousandsSeparator = $null
  try { $origUseSystemSeparators = $excel.UseSystemSeparators } catch {}
  try { $origDecimalSeparator = [string]$excel.DecimalSeparator } catch {}
  try { $origThousandsSeparator = [string]$excel.ThousandsSeparator } catch {}

  # NBSP (U+00A0) used by fr-FR for thousands grouping in many Excel installs.
  $nbsp = [string][char]0x00A0

  $localeConfigs = @(
    [pscustomobject]@{ localeId = "en-US"; decimalSeparator = "."; thousandsSeparator = "," },
    [pscustomobject]@{ localeId = "de-DE"; decimalSeparator = ","; thousandsSeparator = "." },
    [pscustomobject]@{ localeId = "fr-FR"; decimalSeparator = ","; thousandsSeparator = $nbsp },
    # `es-ES` shares the same punctuation as `de-DE`, but has different function translations.
    [pscustomobject]@{ localeId = "es-ES"; decimalSeparator = ","; thousandsSeparator = "." }
  )

  $cases = @(
    [pscustomobject]@{ label = "scalar:1234.56"; inputFormula = "=1234.56" },
    [pscustomobject]@{ label = "scalar:1234567"; inputFormula = "=1234567" },
    [pscustomobject]@{ label = "scalar:1.23E3"; inputFormula = "=1.23E3" },
    [pscustomobject]@{ label = "call:SUM(1234.56,0.5)"; inputFormula = "=SUM(1234.56,0.5)" }
  )

  $localeResults = New-Object System.Collections.Generic.List[object]

  foreach ($cfg in $localeConfigs) {
    # Configure numeric separators. This affects how Excel localizes numeric literals and can
    # influence the argument separator it chooses in `FormulaLocal` (to avoid conflicts with the
    # decimal separator).
    try { $excel.UseSystemSeparators = $false } catch {}
    try { $excel.DecimalSeparator = $cfg.decimalSeparator } catch {}
    try { $excel.ThousandsSeparator = $cfg.thousandsSeparator } catch {}

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "Sheet1"

    $caseResults = New-Object System.Collections.Generic.List[object]

    for ($i = 0; $i -lt $cases.Count; $i++) {
      $row = $i + 1
      $cell = $sheet.Cells.Item($row, 1) # Column A
      $case = $cases[$i]

      Set-RangeFormula -RangeObj $cell -Formula $case.inputFormula

      $formula = [string]$cell.Formula
      $formulaLocal = [string]$cell.FormulaLocal
      $groupedCandidate = Add-GroupingToFormulaLocal `
        -FormulaLocal $formulaLocal `
        -DecimalSeparator $cfg.decimalSeparator `
        -ThousandsSeparator $cfg.thousandsSeparator

      $afterLocal = $null
      $afterFormula = $null
      $groupedSetError = $null
      try {
        $cell.FormulaLocal = $groupedCandidate
        $afterLocal = [string]$cell.FormulaLocal
        $afterFormula = [string]$cell.Formula
      } catch {
        $groupedSetError = $_.Exception.Message
      }

      $caseResults.Add([pscustomobject]@{
        label = $case.label
        inputFormula = $case.inputFormula
        formula = $formula
        formulaLocal = $formulaLocal
        groupedFormulaLocalCandidate = $groupedCandidate
        formulaLocalAfterGroupedSet = $afterLocal
        formulaAfterGroupedSet = $afterFormula
        groupedSetError = $groupedSetError
      }) | Out-Null
    }

    $localeResults.Add([pscustomobject]@{
      localeId = $cfg.localeId
      requestedSeparators = [pscustomobject]@{
        decimalSeparator = $cfg.decimalSeparator
        thousandsSeparator = $cfg.thousandsSeparator
      }
      effectiveSeparators = [pscustomobject]@{
        useSystemSeparators = $excel.UseSystemSeparators
        decimalSeparator = [string]$excel.DecimalSeparator
        thousandsSeparator = [string]$excel.ThousandsSeparator
      }
      cases = $caseResults
    }) | Out-Null

    # Close workbook before switching separators again.
    $workbook.Close($false) | Out-Null
    Release-ComObject $sheet
    Release-ComObject $workbook
    $sheet = $null
    $workbook = $null
  }

  # Best-effort Excel UI locale id (LCID). The msoLanguageIDUI constant is 2.
  $uiLanguage = $null
  try { $uiLanguage = $excel.LanguageSettings.LanguageID(2) } catch { $uiLanguage = $null }

  $out = [pscustomobject]@{
    source = "Excel COM FormulaLocal numeric literal formatting probe"
    excelVersion = [string]$excel.Version
    excelUiLanguageId = $uiLanguage
    locales = $localeResults
  }

  $out | ConvertTo-Json -Depth 12 | Out-File -FilePath $outFile -Encoding UTF8
  Write-Host "Wrote $outFile"
} finally {
  try {
    if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
  } catch {}

  try {
    # Restore separators so we don't leave a running Excel instance in a surprising state.
    if ($null -ne $origUseSystemSeparators) { $excel.UseSystemSeparators = $origUseSystemSeparators }
    if ($null -ne $origDecimalSeparator) { $excel.DecimalSeparator = $origDecimalSeparator }
    if ($null -ne $origThousandsSeparator) { $excel.ThousandsSeparator = $origThousandsSeparator }
  } catch {}

  try {
    if ($null -ne $excel) { $excel.Quit() | Out-Null }
  } catch {}

  Release-ComObject $sheet
  Release-ComObject $workbook
  Release-ComObject $excel

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

