<#
.SYNOPSIS
  Evaluate a corpus of formulas in real Microsoft Excel via COM automation.

.DESCRIPTION
  Reads a case corpus JSON (tests/compatibility/excel-oracle/cases.json),
  executes each case in an in-memory workbook, and writes a machine-readable
  JSON dataset of expected results ("Excel oracle").

  This is Windows-only and requires Microsoft Excel desktop installed.

.PARAMETER CasesPath
  Path to cases.json

.PARAMETER OutPath
  Path where the Excel oracle dataset JSON will be written.

.PARAMETER MaxCases
  Optional cap for debugging (run only the first N cases).

.PARAMETER IncludeTags
  Optional list of case tags to include. If provided, only cases that contain
  at least one of these tags are evaluated.

.PARAMETER ExcludeTags
  Optional list of case tags to exclude. Any case containing one of these tags
  is skipped.

.PARAMETER Visible
  Make Excel visible while running (useful for debugging).

.PARAMETER DryRun
  Print a summary of how many cases would be evaluated (after tag filtering / MaxCases) and exit without running Excel.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$CasesPath,

  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [int]$MaxCases = 0,

  [string[]]$IncludeTags = @(),

  [string[]]$ExcludeTags = @(),

  [switch]$Visible,

  [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

  # Prefer `Formula2` when available so dynamic array formulas behave the same
  # way users see them in modern Excel (avoids implicit-intersection `@` quirks).
  try {
    $RangeObj.Formula2 = $Formula
  } catch {
    $RangeObj.Formula = $Formula
  }
}

function Encode-CellValue {
  param([object]$CellRange)

  $v = $CellRange.Value2
  if ($null -eq $v) {
    return [ordered]@{ t = "blank" }
  }

  # PowerShell sometimes returns Excel errors as Int32 error codes; .Text is the
  # most reliable way to determine error strings (#DIV/0!, #N/A, #SPILL!, ...).
  $text = $CellRange.Text
  if ($text -is [string] -and $text.StartsWith("#") -and -not ($v -is [string])) {
    return [ordered]@{ t = "e"; v = $text }
  }

  if ($v -is [bool]) {
    return [ordered]@{ t = "b"; v = [bool]$v }
  }

  if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) {
    return [ordered]@{ t = "n"; v = [double]$v }
  }

  return [ordered]@{ t = "s"; v = [string]$v }
}

function Encode-RangeValue {
  param([object]$RangeObj)

  $rows = $RangeObj.Rows.Count
  $cols = $RangeObj.Columns.Count

  if ($rows -eq 1 -and $cols -eq 1) {
    $encoded = Encode-CellValue -CellRange $RangeObj
    return [ordered]@{
      value = $encoded
      address = $RangeObj.Address($false, $false)
      displayText = [string]$RangeObj.Text
    }
  }

  $outRows = @()
  for ($r = 1; $r -le $rows; $r++) {
    $row = @()
    for ($c = 1; $c -le $cols; $c++) {
      $cell = $RangeObj.Item($r, $c)
      try {
        $row += ,(Encode-CellValue -CellRange $cell)
      } finally {
        Release-ComObject $cell
      }
    }
    $outRows += ,$row
  }

  $topLeft = $RangeObj.Item(1, 1)
  try {
    $display = [string]$topLeft.Text
  } finally {
    Release-ComObject $topLeft
  }

  return [ordered]@{
    value = [ordered]@{ t = "arr"; rows = $outRows }
    address = $RangeObj.Address($false, $false)
    displayText = $display
  }
}

if (-not (Test-Path -LiteralPath $CasesPath)) {
  throw "CasesPath not found: $CasesPath"
}

$casesJson = Get-Content -LiteralPath $CasesPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($casesJson.schemaVersion -ne 1) {
  throw "Unsupported cases schemaVersion: $($casesJson.schemaVersion)"
}

$caseList = @($casesJson.cases)

$include = @($IncludeTags | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
$exclude = @($ExcludeTags | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })

if ($include.Count -gt 0 -or $exclude.Count -gt 0) {
  $caseList = $caseList | Where-Object {
    $tags = @()
    if ($null -ne $_.tags) { $tags = @($_.tags | ForEach-Object { [string]$_ }) }

    if ($include.Count -gt 0) {
      $matched = $false
      foreach ($t in $tags) {
        if ($include -contains $t) { $matched = $true; break }
      }
      if (-not $matched) { return $false }
    }

    if ($exclude.Count -gt 0) {
      foreach ($t in $tags) {
        if ($exclude -contains $t) { return $false }
      }
    }

    return $true
  }
}

if ($MaxCases -gt 0) {
  $caseList = $caseList | Select-Object -First $MaxCases
}

$caseHash = (Get-FileHash -LiteralPath $CasesPath -Algorithm SHA256).Hash.ToLowerInvariant()

if ($DryRun) {
  Write-Host "Dry run: run-excel-oracle.ps1"
  Write-Host ""
  Write-Host "CasesPath: $CasesPath"
  Write-Host "OutPath:   $OutPath"
  Write-Host "Cases selected: $($caseList.Count)"
  Write-Host "cases.json sha256: $caseHash"
  if ($include.Count -gt 0) {
    Write-Host "IncludeTags: $($include -join ',')"
  } else {
    Write-Host "IncludeTags: <none>"
  }
  if ($exclude.Count -gt 0) {
    Write-Host "ExcludeTags: $($exclude -join ',')"
  } else {
    Write-Host "ExcludeTags: <none>"
  }
  Write-Host "Visible: $([bool]$Visible)"
  Write-Host ""
  Write-Host "No files were written; Excel was not started."
  return
}

$excel = $null
$workbook = $null
$sheet = $null
$origUseSystemSeparators = $null
$origDecimalSeparator = $null
$origThousandsSeparator = $null

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

  # For deterministic text->number coercion (e.g. "1,234", "$1,234.50"), force US-style
  # decimal/thousands separators regardless of the runner machine locale.
  #
  # This makes the oracle dataset more portable across self-hosted runners.
  try {
    $origUseSystemSeparators = $excel.UseSystemSeparators
    $origDecimalSeparator = $excel.DecimalSeparator
    $origThousandsSeparator = $excel.ThousandsSeparator

    $excel.UseSystemSeparators = $false
    $excel.DecimalSeparator = "."
    $excel.ThousandsSeparator = ","
  } catch {
    # Best-effort; Excel/Office versions can differ in COM surface area.
  }

  # Manual calculation for performance; we explicitly calculate after setting inputs.
  # xlCalculationManual = -4135, xlCalculationAutomatic = -4105
  $excel.Calculation = -4135

  $workbook = $excel.Workbooks.Add()
  # Ensure a deterministic date system (Excel for Windows defaults to 1900,
  # but this can be toggled per-workbook).
  try { $workbook.Date1904 = $false } catch {}
  $sheet = $workbook.Worksheets.Item(1)
  $sheet.Name = "Oracle"

  $results = New-Object System.Collections.Generic.List[object]

  $i = 0
  foreach ($case in $caseList) {
    $i++
    Write-Verbose ("[{0}/{1}] {2}" -f $i, $caseList.Count, $case.id)

    $sheet.Cells.Clear()

    # Apply inputs
    $inputs = @()
    if ($null -ne $case.inputs) { $inputs = @($case.inputs) }
    foreach ($input in $inputs) {
      $cellRef = [string]$input.cell
      $range = $sheet.Range($cellRef)
      try {
        if ($null -ne $input.formula) {
          Set-RangeFormula -RangeObj $range -Formula ([string]$input.formula)
        } elseif ($null -eq $input.value) {
          $range.ClearContents()
        } else {
          $range.Value2 = $input.value
        }
      } finally {
        Release-ComObject $range
      }
    }

    # Apply formula under test
    $outputCell = if ($null -ne $case.outputCell) { [string]$case.outputCell } else { "C1" }
    $formulaRange = $sheet.Range($outputCell)
    try {
      Set-RangeFormula -RangeObj $formulaRange -Formula ([string]$case.formula)
    } finally {
      Release-ComObject $formulaRange
    }

    # Calculate
    $excel.Calculate()

    # Read result (support spill where available)
    $resultStart = $sheet.Range($outputCell)
    $resultRange = $resultStart
    try {
      try {
        $spill = $resultStart.SpillingToRange
        if ($null -ne $spill) {
          $resultRange = $spill
        }
      } catch {
        # Property not available or not a spilling formula. Ignore.
      }

      $encoded = Encode-RangeValue -RangeObj $resultRange

      $results.Add([ordered]@{
        caseId = $case.id
        outputCell = $outputCell
        result = $encoded.value
        address = $encoded.address
        displayText = $encoded.displayText
      }) | Out-Null
    } finally {
      if ($resultRange -ne $resultStart) { Release-ComObject $resultRange }
      Release-ComObject $resultStart
    }
  }

  $payload = [ordered]@{
    schemaVersion = 1
    generatedAt = (Get-Date).ToUniversalTime().ToString("o")
    source = [ordered]@{
      kind = "excel"
      version = [string]$excel.Version
      build = [string]$excel.Build
      operatingSystem = [string]$excel.OperatingSystem
    }
    caseSet = [ordered]@{
      path = $CasesPath
      sha256 = $caseHash
      count = $caseList.Count
    }
    results = $results
  }

  $outDir = Split-Path -Parent $OutPath
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
  }

  $json = $payload | ConvertTo-Json -Depth 50
  $fullOutPath = [System.IO.Path]::GetFullPath($OutPath)
  [System.IO.File]::WriteAllText($fullOutPath, $json + "`n", [System.Text.UTF8Encoding]::new($false))
} finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($false) } catch {}
  }
  if ($null -ne $excel) {
    # Restore locale/separator settings before quitting (best-effort).
    try {
      if ($null -ne $origUseSystemSeparators) {
        $excel.UseSystemSeparators = $origUseSystemSeparators
      }
      if ($null -ne $origDecimalSeparator) {
        $excel.DecimalSeparator = $origDecimalSeparator
      }
      if ($null -ne $origThousandsSeparator) {
        $excel.ThousandsSeparator = $origThousandsSeparator
      }
    } catch {}
    try { $excel.Quit() } catch {}
  }

  Release-ComObject $sheet
  Release-ComObject $workbook
  Release-ComObject $excel

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
