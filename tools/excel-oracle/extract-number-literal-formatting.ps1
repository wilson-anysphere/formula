<#
.SYNOPSIS
  Probe how Excel serializes numeric literals in `FormulaLocal` via COM automation.

.DESCRIPTION
  This script writes a handful of canonical (en-US) formulas containing numeric literals using
  `Range.Formula2` and then records what Excel reports through `Range.FormulaLocal`.

  This is primarily used as an "oracle" to validate `crates/formula-engine`'s
  `locale::localize_formula` behavior for:
    - decimal separator localization (e.g. `1234.56` -> `1234,56` in de-DE)
    - whether Excel inserts thousands/grouping separators in formulas
    - whether Excel preserves leading zeros in numeric literals (e.g. `0001`)
    - scientific notation behavior (e.g. `1E3`)

  NOTE: Excel's UI language cannot be reliably switched via COM automation. `-LocaleId` is used
  only for output labeling. Ensure Excel is configured to the desired UI language before running.

.PARAMETER LocaleId
  Locale id used for output naming / labeling (e.g. de-DE, es-ES, fr-FR).

.PARAMETER OutPath
  Output path for the generated JSON.

  - If OutPath ends with ".json", it is treated as a file path.
  - Otherwise, it is treated as a directory, and "<LocaleId>.number-literal-formatting.json" is written inside it.

.PARAMETER Visible
  Make Excel visible while running (useful for debugging).

.EXAMPLE
  # From repo root on Windows (requires Excel desktop installed)
  powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-number-literal-formatting.ps1 `
    -LocaleId de-DE `
    -OutPath out/
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$LocaleId,

  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
  throw "extract-number-literal-formatting.ps1 is Windows-only (requires Excel desktop + COM automation)."
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
  param(
    [Parameter(Mandatory = $true)][string]$LocaleId,
    [Parameter(Mandatory = $true)][string]$OutPath
  )

  $full = Resolve-FullPath -Path $OutPath
  if ($full.ToLowerInvariant().EndsWith(".json")) {
    $dir = Split-Path -Parent $full
    if ($dir -and !(Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    return $full
  }

  if (!(Test-Path $full)) { [void](New-Item -ItemType Directory -Path $full -Force) }
  return (Join-Path $full "$LocaleId.number-literal-formatting.json")
}

function Set-RangeFormula {
  param(
    [Parameter(Mandatory = $true)][object]$RangeObj,
    [Parameter(Mandatory = $true)][string]$Formula
  )

  # Prefer `Formula2` when available (recorded output still reads from `FormulaLocal`).
  try {
    $RangeObj.Formula2 = $Formula
  } catch {
    $RangeObj.Formula = $Formula
  }
}

$outFile = Ensure-OutputFile -LocaleId $LocaleId -OutPath $OutPath

$excel = $null
$workbook = $null
$sheet = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = [bool]$Visible
  $excel.DisplayAlerts = $false

  $workbook = $excel.Workbooks.Add()
  $sheet = $workbook.Worksheets.Item(1)
  $sheet.Name = "Sheet1"

  $cases = @(
    @{ label = "1234.56"; formula = "=SUM(1234.56,0.5)" },
    @{ label = "1234567.89"; formula = "=SUM(1234567.89,0.5)" },
    @{ label = "1000"; formula = "=SUM(1000,0)" },
    @{ label = "leading zeros"; formula = "=SUM(0001,0)" },
    @{ label = "scientific notation"; formula = "=SUM(1E3,0)" }
  )

  $results = New-Object System.Collections.Generic.List[object]
  for ($i = 0; $i -lt $cases.Count; $i++) {
    $row = $i + 1
    $cell = $sheet.Cells.Item($row, 4) # Column D
    $case = $cases[$i]
    Set-RangeFormula -RangeObj $cell -Formula $case.formula

    $results.Add([pscustomobject]@{
      label = $case.label
      inputFormula = $case.formula
      formula = [string]$cell.Formula
      formulaLocal = [string]$cell.FormulaLocal
    }) | Out-Null
  }

  # Best-effort Excel UI locale id (LCID). The msoLanguageIDUI constant is 2.
  $uiLanguage = $null
  try {
    $uiLanguage = $excel.LanguageSettings.LanguageID(2)
  } catch {
    $uiLanguage = $null
  }

  # Excel.Application.International constants (numeric ids):
  # - xlDecimalSeparator = 3
  # - xlThousandsSeparator = 4
  # - xlListSeparator = 5
  $intl = $null
  try {
    $intl = [pscustomobject]@{
      decimalSeparator = [string]$excel.International(3)
      thousandsSeparator = [string]$excel.International(4)
      listSeparator = [string]$excel.International(5)
    }
  } catch {
    $intl = $null
  }

  $out = [pscustomobject]@{
    source = "Excel COM FormulaLocal number literal probe"
    localeId = $LocaleId
    excelUiLanguageId = $uiLanguage
    excelInternational = $intl
    cases = $results
  }

  $json = $out | ConvertTo-Json -Depth 8
  $json | Out-File -FilePath $outFile -Encoding UTF8
  Write-Host "Wrote $outFile"
} finally {
  try {
    if ($null -ne $workbook) { $workbook.Close($false) | Out-Null }
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

