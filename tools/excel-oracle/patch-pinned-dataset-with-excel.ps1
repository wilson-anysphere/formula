<#
.SYNOPSIS
  Patch the pinned Excel-oracle dataset with results from real Microsoft Excel.

.DESCRIPTION
  This script is a convenience wrapper for a common workflow:

    1) Run `tools/excel-oracle/run-excel-oracle.ps1` on a small *subset* corpus to capture
       expected results in real Excel (COM automation).
    2) Merge those results into the committed pinned dataset
       (`tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json`) by overwriting the
       matching `caseId`s (without touching the rest of the corpus).

  This makes it easy to gradually replace the synthetic CI baseline with real Excel results for
  targeted edge cases (like odd-coupon bonds) without regenerating the entire dataset.

  Note: the subset corpus should reuse canonical `caseId`s from `tests/compatibility/excel-oracle/cases.json`.

.PARAMETER SubsetCasesPath
  Path to a subset corpus (default: tools/excel-oracle/odd_coupon_long_stub_cases.json).
  Other useful subsets include:
    - tools/excel-oracle/odd_coupon_boundary_cases.json (odd-coupon date boundary scenarios)
    - tools/excel-oracle/odd_coupon_validation_cases.json (negative yields / yield-domain edges / negative rate)

.PARAMETER CasesPath
  Path to the canonical cases.json corpus (default: tests/compatibility/excel-oracle/cases.json).

.PARAMETER PinnedDatasetPath
  Path to the pinned dataset to patch (default: tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json).

.PARAMETER ExcelOutPath
  Where to write the temporary Excel results JSON. If omitted, a temporary file is used.

.PARAMETER MaxCases
  Optional cap for debugging (passed through to run-excel-oracle.ps1).

.PARAMETER IncludeTags
  Optional tags to include (passed through to run-excel-oracle.ps1). This can be used instead of
  a subset corpus by pointing SubsetCasesPath at the canonical cases.json and filtering by tag.

.PARAMETER ExcludeTags
  Optional tags to exclude (passed through to run-excel-oracle.ps1).

.PARAMETER Visible
  Make Excel visible while running (passed through to run-excel-oracle.ps1).
#>

[CmdletBinding()]
param(
  [string]$SubsetCasesPath = "tools/excel-oracle/odd_coupon_long_stub_cases.json",
  [string]$CasesPath = "tests/compatibility/excel-oracle/cases.json",
  [string]$PinnedDatasetPath = "tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json",
  [string]$ExcelOutPath = "",
  [int]$MaxCases = 0,
  [string[]]$IncludeTags = @(),
  [string[]]$ExcludeTags = @(),
  [switch]$Visible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "../..")

$subsetCasesFull = if ([System.IO.Path]::IsPathRooted($SubsetCasesPath)) { $SubsetCasesPath } else { Join-Path $repoRoot $SubsetCasesPath }
$casesFull = if ([System.IO.Path]::IsPathRooted($CasesPath)) { $CasesPath } else { Join-Path $repoRoot $CasesPath }
$pinnedFull = if ([System.IO.Path]::IsPathRooted($PinnedDatasetPath)) { $PinnedDatasetPath } else { Join-Path $repoRoot $PinnedDatasetPath }

$excelResultsFull = $ExcelOutPath
$cleanupExcelResults = $false
if ([string]::IsNullOrWhiteSpace($excelResultsFull)) {
  $excelResultsFull = Join-Path ([System.IO.Path]::GetTempPath()) ("excel-oracle-results-" + [guid]::NewGuid().ToString("n") + ".json")
  $cleanupExcelResults = $true
}
if (-not [System.IO.Path]::IsPathRooted($excelResultsFull)) {
  $excelResultsFull = Join-Path $repoRoot $excelResultsFull
}

$excelScript = Join-Path $repoRoot "tools/excel-oracle/run-excel-oracle.ps1"
$updateScript = Join-Path $repoRoot "tools/excel-oracle/update_pinned_dataset.py"

if (-not (Test-Path -LiteralPath $subsetCasesFull)) {
  throw "SubsetCasesPath not found: $subsetCasesFull"
}
if (-not (Test-Path -LiteralPath $casesFull)) {
  throw "CasesPath not found: $casesFull"
}
if (-not (Test-Path -LiteralPath $pinnedFull)) {
  throw "PinnedDatasetPath not found: $pinnedFull"
}

Push-Location $repoRoot
try {
  Write-Host "Generating Excel results from subset corpus: $subsetCasesFull"
  & $excelScript -CasesPath $subsetCasesFull -OutPath $excelResultsFull -Visible:$Visible -MaxCases $MaxCases -IncludeTags $IncludeTags -ExcludeTags $ExcludeTags

  Write-Host ""
  Write-Host "Patching pinned dataset (overwrite existing case IDs) -> $pinnedFull"
  python $updateScript `
    --cases $casesFull `
    --pinned $pinnedFull `
    --merge-results $excelResultsFull `
    --overwrite-existing `
    --no-engine

  Write-Host ""
  Write-Host "Done. Review and commit the updated pinned dataset:"
  Write-Host "  git diff -- tests/compatibility/excel-oracle/datasets/excel-oracle.pinned.json"
} finally {
  Pop-Location
  if ($cleanupExcelResults -and (Test-Path -LiteralPath $excelResultsFull)) {
    Remove-Item -LiteralPath $excelResultsFull -Force -ErrorAction SilentlyContinue
  }
}
