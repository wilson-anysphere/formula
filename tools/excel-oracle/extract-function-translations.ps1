<#
.SYNOPSIS
  Extract localized Excel function identifiers for a given locale via COM automation.

.DESCRIPTION
  This script uses a real Microsoft Excel desktop installation to round-trip
  canonical (en-US) formulas through `Range.Formula` / `Range.FormulaLocal` and
  derive a mapping from canonical function identifiers to their localized names.

  The output JSON is suitable for consumption by:

    node scripts/generate-locale-function-tsv.js

  NOTE: Excel's displayed function names depend on the Office language packs /
  editing language configuration of the installed Excel. This script does not
  (and cannot reliably) switch Excel's UI language; `-LocaleId` is used only
  for naming and metadata in the output file.

.PARAMETER LocaleId
  Locale id used for output naming / labeling (e.g. de-DE, es-ES, fr-FR).

.PARAMETER OutPath
  Output path for the generated JSON.

  - If OutPath ends with ".json", it is treated as a file path.
  - Otherwise, it is treated as a directory, and "<LocaleId>.json" is written inside it.

.PARAMETER Visible
  Make Excel visible while running (useful for debugging).

.PARAMETER MaxFunctions
  Optional cap for debugging (extract only the first N catalog functions).
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$LocaleId,

  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [switch]$Visible,

  [int]$MaxFunctions = 0
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

function Resolve-FullPath {
  param([Parameter(Mandatory = $true)][string]$Path)
  if ([System.IO.Path]::IsPathRooted($Path)) {
    return [System.IO.Path]::GetFullPath($Path)
  }
  return [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Path))
}

function Placeholder-ForArgType {
  param([Parameter(Mandatory = $true)][string]$ArgType)
  switch ($ArgType) {
    "bool" { return "TRUE" }
    "text" { return '"x"' }
    "number" { return "1" }
    "any" { return "1" }
    default { return "1" }
  }
}

function Build-MinimalFormula {
  param(
    [Parameter(Mandatory = $true)][string]$FunctionName,
    [Parameter(Mandatory = $true)][int]$MinArgs,
    [Parameter(Mandatory = $true)][object[]]$ArgTypes
  )

  # Some functions have syntactic constraints that aren't captured by the
  # catalog's `arg_types` (e.g. LET requires a name token as its first arg).
  # Special-case them so we can still extract translations when Excel supports
  # the function.
  if ($FunctionName -eq "LET") {
    return "=LET(x,1,x)"
  }

  # LAMBDA() is only useful when invoked; using an invocation keeps the formula
  # syntactically valid and avoids edge-cases where Excel treats a bare LAMBDA
  # value as an error.
  if ($FunctionName -eq "LAMBDA") {
    return "=LAMBDA(x,x)(1)"
  }

  # Functions that require a LAMBDA argument often reject formulas that use a
  # scalar placeholder (e.g. BYROW(1,1)). Use a minimal lambda instead.
  if ($FunctionName -eq "BYROW") {
    return "=BYROW(1,LAMBDA(x,x))"
  }

  if ($FunctionName -eq "BYCOL") {
    return "=BYCOL(1,LAMBDA(x,x))"
  }

  if ($FunctionName -eq "MAP") {
    return "=MAP(1,LAMBDA(x,x))"
  }

  if ($FunctionName -eq "MAKEARRAY") {
    return "=MAKEARRAY(1,1,LAMBDA(r,c,1))"
  }

  if ($FunctionName -eq "REDUCE") {
    return "=REDUCE(0,1,LAMBDA(a,b,a+b))"
  }

  if ($FunctionName -eq "SCAN") {
    return "=SCAN(0,1,LAMBDA(a,b,a+b))"
  }

  if ($MinArgs -le 0) {
    return "=$FunctionName()"
  }

  $args = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $MinArgs; $i++) {
    $t = $null
    if ($null -ne $ArgTypes -and $ArgTypes.Count -gt 0) {
      if ($i -lt $ArgTypes.Count) {
        $t = [string]$ArgTypes[$i]
      } else {
        # Some catalog entries provide a single repeated type for variadic functions.
        $t = [string]$ArgTypes[$ArgTypes.Count - 1]
      }
    } else {
      $t = "any"
    }
    $args.Add((Placeholder-ForArgType -ArgType $t)) | Out-Null
  }

  return "=$FunctionName(" + ($args -join ",") + ")"
}

function Parse-LocalizedFunctionName {
  param([Parameter(Mandatory = $true)][string]$FormulaLocal)

  $s = $FormulaLocal.Trim()
  if ($s.StartsWith("=")) {
    $s = $s.Substring(1)
  }

  # Strip leading implicit-intersection marker if present.
  if ($s.StartsWith("@")) {
    $s = $s.Substring(1)
  }

  # Some older Excel builds prefix unknown/new functions with `_xlfn.`.
  if ($s.StartsWith("_xlfn.")) {
    $s = $s.Substring(6)
  }

  # Defensive: Excel occasionally emits formulas with a leading '+'.
  if ($s.StartsWith("+")) {
    $s = $s.Substring(1)
  }

  $idx = $s.IndexOf("(")
  if ($idx -ge 0) {
    return $s.Substring(0, $idx).Trim()
  }

  return $s.Trim()
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot ".." "..")
$catalogPath = Join-Path $repoRoot "shared" "functionCatalog.json"

if (-not (Test-Path -LiteralPath $catalogPath)) {
  throw "Function catalog not found: $catalogPath"
}

$catalog = Get-Content -LiteralPath $catalogPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($null -eq $catalog.functions) {
  throw "Invalid function catalog shape: missing .functions in $catalogPath"
}

$functionList = @($catalog.functions)
if ($MaxFunctions -gt 0) {
  $functionList = $functionList | Select-Object -First $MaxFunctions
}

$fullOutPath = Resolve-FullPath -Path $OutPath
$outFilePath = $fullOutPath
$outDir = $null

if ($fullOutPath.EndsWith(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
  $outFilePath = $fullOutPath
  $outDir = Split-Path -Parent $outFilePath
} else {
  $outDir = $fullOutPath
  $outFilePath = Join-Path $outDir "$LocaleId.json"
}

if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

$excel = $null
$workbook = $null
$sheet = $null
$cell = $null
$origCalculation = $null

$translations = @{}
$skipped = New-Object System.Collections.Generic.List[string]

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

  # Avoid evaluating formulas during extraction (e.g. WEBSERVICE/CUBE functions).
  # We only need Excel to parse the formula so we can read back FormulaLocal.
  # xlCalculationManual = -4135, xlCalculationAutomatic = -4105
  try {
    $origCalculation = $excel.Calculation
    $excel.Calculation = -4135
  } catch {
    # Best-effort; ignore if not supported by this Excel/Office build.
  }

  $workbook = $excel.Workbooks.Add()
  $sheet = $workbook.Worksheets.Item(1)
  $sheet.Name = "FunctionTranslations"
  $cell = $sheet.Range("A1")

  $i = 0
  foreach ($fn in $functionList) {
    $i++
    $canonicalName = [string]$fn.name
    $minArgs = [int]$fn.min_args
    $argTypes = @()
    if ($null -ne $fn.arg_types) { $argTypes = @($fn.arg_types | ForEach-Object { [string]$_ }) }

    $formula = Build-MinimalFormula -FunctionName $canonicalName -MinArgs $minArgs -ArgTypes $argTypes

    try {
      $cell.Clear()
      Set-RangeFormula -RangeObj $cell -Formula $formula
      $local = [string]$cell.FormulaLocal
      if (-not $local) {
        throw "FormulaLocal was empty"
      }
      $localizedName = Parse-LocalizedFunctionName -FormulaLocal $local
      if (-not $localizedName) {
        throw "Failed to parse localized name from FormulaLocal: $local"
      }

      $translations[$canonicalName] = $localizedName
    } catch {
      $skipped.Add($canonicalName) | Out-Null
      Write-Warning ("[{0}/{1}] Skipping {2}: {3}" -f $i, $functionList.Count, $canonicalName, $_.Exception.Message)
      continue
    }
  }

  $orderedTranslations = [ordered]@{}
  foreach ($k in ($translations.Keys | Sort-Object)) {
    $orderedTranslations[$k] = [string]$translations[$k]
  }

  $payload = [ordered]@{
    source = "Microsoft Excel ($LocaleId) function name translations via Range.Formula/FormulaLocal round-trip (generated by tools/excel-oracle/extract-function-translations.ps1)."
    translations = $orderedTranslations
  }

  $json = $payload | ConvertTo-Json -Depth 10
  [System.IO.File]::WriteAllText($outFilePath, $json + "`n", [System.Text.UTF8Encoding]::new($false))

  Write-Host "Wrote $($orderedTranslations.Count) translations to: $outFilePath"
  if ($skipped.Count -gt 0) {
    Write-Warning ("Skipped {0} functions rejected by Excel: {1}" -f $skipped.Count, ($skipped -join ", "))
  }
} finally {
  if ($null -ne $workbook) {
    try { $workbook.Close($false) } catch {}
  }
  if ($null -ne $excel) {
    try {
      if ($null -ne $origCalculation) {
        $excel.Calculation = $origCalculation
      }
    } catch {}
    try { $excel.Quit() } catch {}
  }

  Release-ComObject $cell
  Release-ComObject $sheet
  Release-ComObject $workbook
  Release-ComObject $excel

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
