<#
.SYNOPSIS
  Probe whether Excel localizes table structured-reference item keywords (e.g. `#Headers`) via COM automation.

.DESCRIPTION
  Excel has a small set of reserved "item keywords" that can appear inside table structured references:

    - `[#All]`, `[#Data]`, `[#Headers]`, `[#Totals]`, `[#This Row]` (and `@` / `[@Col]`)

  This script creates a small workbook with a single table named `Table1` and then round-trips a few
  canonical (en-US) formulas through `Range.Formula` / `Range.FormulaLocal`.

  The output JSON captures both the canonical formula string and whatever Excel reports via
  `FormulaLocal` for the current Excel UI language configuration.

  NOTE: Like `extract-function-translations.ps1`, this script cannot reliably switch Excel's UI language;
  `-LocaleId` is used only for naming/metadata in the output. Ensure Excel is configured to the desired UI
  language before running.

.PARAMETER LocaleId
  Locale id used for output naming / labeling (e.g. de-DE, es-ES, fr-FR).

.PARAMETER OutPath
  Output path for the generated JSON.

  - If OutPath ends with ".json", it is treated as a file path.
  - Otherwise, it is treated as a directory, and "<LocaleId>.structured-ref-keywords.json" is written inside it.

.PARAMETER Visible
  Make Excel visible while running (useful for debugging).

.EXAMPLE
  # From repo root on Windows (requires Excel desktop installed)
  powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-structured-reference-keywords.ps1 `
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
  throw "extract-structured-reference-keywords.ps1 is Windows-only (requires Excel desktop + COM automation)."
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
  return (Join-Path $full "$LocaleId.structured-ref-keywords.json")
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

$outFile = Ensure-OutputFile -LocaleId $LocaleId -OutPath $OutPath

$excel = $null
$workbook = $null
$sheet = $null
$table = $null

try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = [bool]$Visible
  $excel.DisplayAlerts = $false

  $workbook = $excel.Workbooks.Add()
  $sheet = $workbook.Worksheets.Item(1)
  $sheet.Name = "Sheet1"

  # Build a small table so structured refs resolve in a realistic way.
  $sheet.Range("A1").Value2 = "Qty"
  $sheet.Range("B1").Value2 = "Amount"
  $sheet.Range("A2").Value2 = 1
  $sheet.Range("B2").Value2 = 2
  $sheet.Range("A3").Value2 = 3
  $sheet.Range("B3").Value2 = 4

  # XlListObjectSourceType.xlSrcRange = 1, XlYesNoGuess.xlYes = 1
  $table = $sheet.ListObjects.Add(1, $sheet.Range("A1:B3"), $null, 1)
  $table.Name = "Table1"

  $cases = @(
    @{ label = "#All"; formula = "=SUM(Table1[#All])" },
    @{ label = "#Data"; formula = "=SUM(Table1[#Data])" },
    @{ label = "#Headers"; formula = "=SUM(Table1[#Headers])" },
    @{ label = "#Totals"; formula = "=SUM(Table1[#Totals])" },
    @{ label = "#This Row (item keyword)"; formula = "=SUM(Table1[[#This Row],[Qty]])" },
    @{ label = "@ (this row shorthand)"; formula = "=SUM(Table1[@Qty])" },
    @{ label = "nested separators"; formula = "=SUM(Table1[[#Headers],[Qty]])" }
  )

  $results = New-Object System.Collections.Generic.List[object]
  for ($i = 0; $i -lt $cases.Count; $i++) {
    $row = $i + 1
    $cell = $sheet.Cells.Item($row, 4) # Column D
    $case = $cases[$i]
    Set-RangeFormula -RangeObj $cell -Formula $case.formula

    # Read both canonical and local representations as Excel reports them.
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

  $out = [pscustomobject]@{
    source = "Excel COM FormulaLocal structured reference probe"
    localeId = $LocaleId
    excelUiLanguageId = $uiLanguage
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

  Release-ComObject $table
  Release-ComObject $sheet
  Release-ComObject $workbook
  Release-ComObject $excel

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

