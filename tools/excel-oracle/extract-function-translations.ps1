<#
.SYNOPSIS
  Extract localized Excel function identifiers for a given locale via COM automation.

.DESCRIPTION
  This script uses a real Microsoft Excel desktop installation to round-trip
  canonical (en-US) formulas through `Range.Formula` / `Range.FormulaLocal` and
  derive a mapping from canonical function identifiers to their localized names.

  The output JSON is suitable for consumption by:

    node scripts/generate-locale-function-tsv.js

  For committed locale data, the repo typically normalizes the extracted JSON
  (omitting identity mappings + enforcing stable casing) before regenerating TSVs:

    node scripts/normalize-locale-function-sources.js
    node scripts/generate-locale-function-tsv.js
    node scripts/generate-locale-function-tsv.js --check

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

  Note: this is intended for debugging only; do not commit partial locale sources
  generated with `-MaxFunctions`.

.PARAMETER FailOnSkipped
  Fail the extraction if Excel rejects any functions (i.e. if any canonical functions are
  skipped due to parsing/translation errors or being treated as `_xludf.`).

  This is recommended when generating locale sources intended to be committed to the repo,
  since missing translations silently fall back to identity mappings (English) during TSV
  generation.

.EXAMPLE
  # Generate the de-DE source JSON from a German Excel install (from repo root)
  powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
    -LocaleId de-DE `
    -OutPath crates/formula-engine/src/locale/data/sources/de-DE.json

  # Then normalize + regenerate the committed TSVs:
  node scripts/normalize-locale-function-sources.js
  node scripts/generate-locale-function-tsv.js
  node scripts/generate-locale-function-tsv.js --check

.EXAMPLE
  # Generate the es-ES source JSON from a Spanish Excel install (from repo root)
  powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
    -LocaleId es-ES `
    -OutPath crates/formula-engine/src/locale/data/sources/es-ES.json

  # Then normalize + regenerate the committed TSVs:
  node scripts/normalize-locale-function-sources.js
  node scripts/generate-locale-function-tsv.js
  node scripts/generate-locale-function-tsv.js --check

.EXAMPLE
  # Debug a quick subset while watching Excel + printing per-function formulas
  powershell -ExecutionPolicy Bypass -File tools/excel-oracle/extract-function-translations.ps1 `
    -LocaleId de-DE `
    -OutPath out/de-DE.json `
    -MaxFunctions 50 `
    -Visible `
    -Verbose
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$LocaleId,

  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [switch]$Visible,

  [int]$MaxFunctions = 0,

  [switch]$FailOnSkipped
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
  throw "extract-function-translations.ps1 is Windows-only (requires Excel desktop + COM automation)."
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

  # ISOMITTED() is intended to be used inside LAMBDA to test whether an optional
  # argument was omitted. Some inputs (e.g. `ISOMITTED(1)`) may be rejected by
  # Excel, but using an identifier-like name token keeps the formula valid for
  # parsing/translation even outside LAMBDA.
  if ($FunctionName -eq "ISOMITTED") {
    return "=ISOMITTED(y)"
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

function Expected-SentinelTranslations {
  param([Parameter(Mandatory = $true)][string]$LocaleId)

  switch ($LocaleId) {
    "de-DE" { return @{ SUM = "SUMME"; IF = "WENN"; TRUE = "WAHR"; FALSE = "FALSCH" } }
    "es-ES" { return @{ SUM = "SUMA"; IF = "SI"; TRUE = "VERDADERO"; FALSE = "FALSO" } }
    "fr-FR" { return @{ SUM = "SOMME"; IF = "SI"; TRUE = "VRAI"; FALSE = "FAUX" } }
    default { return $null }
  }
}

function Parse-LocalizedFunctionName {
  param([Parameter(Mandatory = $true)][string]$FormulaLocal)

  $s = $FormulaLocal.Trim()

  # Defensive: Excel sometimes serializes formulas with extra leading markers
  # like `=+SUM(...)` or `=@SUM(...)`. Strip these before inspecting prefixes.
  while ($s.Length -gt 0) {
    $ch = $s.Substring(0, 1)
    if ($ch -eq "=" -or $ch -eq "@" -or $ch -eq "+") {
      $s = $s.Substring(1)
      continue
    }
    break
  }

  # Some older Excel builds prefix unknown/new functions with `_xlfn.`.
  #
  # IMPORTANT: When extracting translations for committed locale sources, treating `_xlfn.` as a
  # normal identifier causes unknown functions to degrade into identity mappings (English), which
  # looks "complete" but is wrong. If we see `_xlfn.`, treat it as unsupported in this Excel build
  # and surface it as a skipped function so callers can retry on a modern Excel 365 install.
  $sawXlfn = $false
  while ($true) {
    if ($s.StartsWith("_xlfn.")) {
      $s = $s.Substring(6)
      $sawXlfn = $true
      continue
    }
    # Some Excel builds use `_xlws.` in compatibility wrappers (commonly nested
    # under `_xlfn.`, e.g. `_xlfn._xlws.WEBSERVICE(...)`).
    if ($s.StartsWith("_xlws.")) {
      $s = $s.Substring(6)
      continue
    }
    break
  }

  if ($sawXlfn) {
    throw "Excel prefixed function with _xlfn. (unsupported/unknown in this Excel build). FormulaLocal=$FormulaLocal"
  }

  # Some Excel builds use `_xludf.` for user-defined / unknown functions.
  # If we see this prefix, Excel did not recognize the function identifier and
  # treated it as a UDF; skip recording a translation so callers can retry with
  # a newer Excel build / correct language pack.
  if ($s.StartsWith("_xludf.")) {
    throw "Excel treated function as user-defined (_xludf.). FormulaLocal=$FormulaLocal"
  }

  $idx = $s.IndexOf("(")
  if ($idx -ge 0) {
    return $s.Substring(0, $idx).Trim()
  }

  return $s.Trim()
}

function Warn-IfExcelLocaleSeemsMisconfigured {
  param(
    [Parameter(Mandatory = $true)][object]$RangeObj,
    [Parameter(Mandatory = $true)][string]$LocaleId
  )

  $expected = Expected-SentinelTranslations -LocaleId $LocaleId
  if ($null -eq $expected) { return }

  $checks = @(
    @{ canonical = "SUM"; formula = "=SUM(1,1)" },
    @{ canonical = "IF"; formula = "=IF(TRUE,1,1)" },
    @{ canonical = "TRUE"; formula = "=TRUE()" },
    @{ canonical = "FALSE"; formula = "=FALSE()" }
  )

  foreach ($c in $checks) {
    $canonical = [string]$c.canonical
    $formula = [string]$c.formula
    $want = [string]$expected[$canonical]

    try {
      $RangeObj.Clear()
      Set-RangeFormula -RangeObj $RangeObj -Formula $formula
      $local = [string]$RangeObj.FormulaLocal
      $got = Parse-LocalizedFunctionName -FormulaLocal $local
      if (-not $got) { throw "Parsed localized name was empty (FormulaLocal=$local)" }

      if ($got.ToUpperInvariant() -ne $want.ToUpperInvariant()) {
        Write-Warning ("Excel locale validation: expected {0} -> {1} for -LocaleId {2}, got {3} (FormulaLocal={4})" -f $canonical, $want, $LocaleId, $got, $local)
      }
    } catch {
      Write-Warning ("Excel locale validation: failed for {0} (formula={1}): {2}" -f $canonical, $formula, $_.Exception.Message)
    }
  }
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

$functionList = @($catalog.functions | Sort-Object -Property name)
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
$excelUiLocale = $null
$excelVersion = $null
$excelBuild = $null

$translations = @{}
$skipped = New-Object System.Collections.Generic.List[string]
$localizedToCanonical = @{}
$duplicates = New-Object System.Collections.Generic.List[string]

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
    $excelUiLocale = $null
  }

  try { $excelVersion = [string]$excel.Version } catch {}
  try { $excelBuild = [string]$excel.Build } catch {}

  Write-Host "Excel: version=$excelVersion build=$excelBuild uiLocale=$excelUiLocale"
  if ($excelUiLocale -and ($excelUiLocale -ne $LocaleId)) {
    Write-Warning "Excel UI locale ($excelUiLocale) does not match -LocaleId ($LocaleId). Ensure Excel is configured for the target locale before extracting."
  }

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

  # If we're extracting for a locale we know well, do a quick sanity check
  # against a few sentinel translations so users don't accidentally generate
  # (e.g.) de-DE.json from an en-US Excel install.
  Warn-IfExcelLocaleSeemsMisconfigured -RangeObj $cell -LocaleId $LocaleId

  $i = 0
  $total = $functionList.Count
  foreach ($fn in $functionList) {
    $i++
    $canonicalName = [string]$fn.name
    Write-Progress -Activity "Extracting function translations" -Status ("{0}/{1} {2}" -f $i, $total, $canonicalName) -PercentComplete ([int](($i / [double]$total) * 100))
    $minArgs = [int]$fn.min_args
    $argTypes = @()
    if ($null -ne $fn.arg_types) { $argTypes = @($fn.arg_types | ForEach-Object { [string]$_ }) }

    $formula = Build-MinimalFormula -FunctionName $canonicalName -MinArgs $minArgs -ArgTypes $argTypes
    Write-Verbose ("{0}: {1}" -f $canonicalName, $formula)

    try {
      $cell.Clear()
      Set-RangeFormula -RangeObj $cell -Formula $formula
      $local = [string]$cell.FormulaLocal
      Write-Verbose ("{0} FormulaLocal: {1}" -f $canonicalName, $local)
      if (-not $local) {
        throw "FormulaLocal was empty"
      }
      $localizedName = Parse-LocalizedFunctionName -FormulaLocal $local
      if (-not $localizedName) {
        throw "Failed to parse localized name from FormulaLocal: $local"
      }
      if ($localizedName.StartsWith("#")) {
        throw "Unexpected FormulaLocal (appears to be an error literal): $local"
      }

      $localizedKey = $localizedName.ToUpperInvariant()
      if ($localizedToCanonical.ContainsKey($localizedKey)) {
        $existing = [string]$localizedToCanonical[$localizedKey]
        if ($existing -ne $canonicalName) {
          $duplicates.Add("$existing,$canonicalName -> $localizedName") | Out-Null
          Write-Warning ("Duplicate localized function name detected: {0} and {1} both map to {2}" -f $existing, $canonicalName, $localizedName)
        }
      } else {
        $localizedToCanonical[$localizedKey] = $canonicalName
      }

      $translations[$canonicalName] = $localizedName
    } catch {
      $skipped.Add($canonicalName) | Out-Null
      Write-Warning ("[{0}/{1}] Skipping {2} (formula={3}): {4}" -f $i, $functionList.Count, $canonicalName, $formula, $_.Exception.Message)
      continue
    }
  }
  Write-Progress -Activity "Extracting function translations" -Completed

  $orderedTranslations = [ordered]@{}
  foreach ($k in ($translations.Keys | Sort-Object)) {
    $orderedTranslations[$k] = [string]$translations[$k]
  }

  if ($FailOnSkipped -and $skipped.Count -gt 0) {
    $skippedSorted = @($skipped | Sort-Object)
    throw ("Extraction failed because Excel rejected {0} functions: {1}`n" -f $skipped.Count, ($skippedSorted -join ", ")) +
      "This usually indicates an unsupported/older Excel build or a missing language pack; " +
      "retry on a modern Excel install configured for the requested locale."
  }

  $payload = [ordered]@{
    # Keep this label stable across Excel updates so running the extractor does not create noisy diffs
    # when only the Office build number changes.
    source = "Microsoft Excel ($LocaleId) function name translations via Range.Formula/FormulaLocal round-trip (generated by tools/excel-oracle/extract-function-translations.ps1)."
    translations = $orderedTranslations
  }

  $json = $payload | ConvertTo-Json -Depth 10
  [System.IO.File]::WriteAllText($outFilePath, $json + "`n", [System.Text.UTF8Encoding]::new($false))

  Write-Host "Wrote $($orderedTranslations.Count) translations to: $outFilePath"
  if ($duplicates.Count -gt 0) {
    $dupsSorted = @($duplicates | Sort-Object)
    Write-Warning ("Detected {0} duplicate localized spellings (generator may fail): {1}" -f $duplicates.Count, ($dupsSorted -join "; "))
  }
  if ($skipped.Count -gt 0) {
    $skippedSorted = @($skipped | Sort-Object)
    Write-Warning ("Skipped {0} functions rejected by Excel: {1}" -f $skipped.Count, ($skippedSorted -join ", "))
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
