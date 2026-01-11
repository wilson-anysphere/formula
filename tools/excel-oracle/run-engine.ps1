<#
.SYNOPSIS
  Run the formula engine against the Excel oracle case corpus.

.DESCRIPTION
  This harness is intentionally engine-agnostic. The expected long-term
  integration is that our formula engine exposes a CLI that can:

    - Read the cases.json corpus (formulas + input cells)
    - Evaluate each case
    - Emit a results JSON file with the same schema as the Excel oracle

  This script is a thin wrapper around that CLI so CI has a stable entrypoint.

.PARAMETER CasesPath
  Path to cases.json

.PARAMETER OutPath
  Path where engine results JSON will be written.

.PARAMETER EngineCommand
  Command line used to invoke the engine.

  If omitted, the script uses $env:FORMULA_ENGINE_CMD.

.PARAMETER MaxCases
  Optional cap for debugging (run only the first N cases).

.PARAMETER IncludeTags
  Optional list of case tags to include. If provided, only cases that contain
  at least one of these tags are evaluated.

.PARAMETER ExcludeTags
  Optional list of case tags to exclude. Any case containing one of these tags
  is skipped.

.NOTES
  If no engine command is provided, this script defaults to running the
  in-repo Rust CLI (`cargo run -p formula-excel-oracle -- ...`).
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

  [string]$EngineCommand
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "../..")
if ([string]::IsNullOrWhiteSpace($env:CARGO_HOME)) {
  $env:CARGO_HOME = Join-Path $repoRoot "target/cargo-home"
}
New-Item -ItemType Directory -Force -Path $env:CARGO_HOME | Out-Null

if (-not $EngineCommand) {
  $EngineCommand = $env:FORMULA_ENGINE_CMD
}

if (-not $EngineCommand) {
  # Default: use the in-repo Rust CLI that evaluates the corpus via formula-engine.
  $cargoArgs = @(
    "run",
    "-p", "formula-excel-oracle",
    "--quiet",
    "--locked",
    "--",
    "--cases", $CasesPath,
    "--out", $OutPath
  )
  if ($MaxCases -gt 0) { $cargoArgs += @("--max-cases", $MaxCases) }
  foreach ($t in $IncludeTags) { if ($t -and $t.Trim() -ne "") { $cargoArgs += @("--include-tag", $t.Trim()) } }
  foreach ($t in $ExcludeTags) { if ($t -and $t.Trim() -ne "") { $cargoArgs += @("--exclude-tag", $t.Trim()) } }

  Write-Host ("Running engine via cargo: cargo {0}" -f ($cargoArgs -join " "))
  & cargo @cargoArgs
  exit $LASTEXITCODE
}

if (-not (Test-Path -LiteralPath $CasesPath)) {
  throw "CasesPath not found: $CasesPath"
}

$outDir = Split-Path -Parent $OutPath
if ($outDir -and -not (Test-Path -LiteralPath $outDir)) {
  New-Item -ItemType Directory -Force -Path $outDir | Out-Null
}

# Convention: engine CLI accepts:
#   --cases <path> --out <path>
#
# This keeps the harness stable while allowing the underlying engine to evolve.
$cmd = "$EngineCommand --cases `"$CasesPath`" --out `"$OutPath`""
if ($MaxCases -gt 0) {
  $cmd = "$cmd --max-cases $MaxCases"
}
foreach ($t in $IncludeTags) {
  if ($t -and $t.Trim() -ne "") {
    $cmd = "$cmd --include-tag `"$($t.Trim())`""
  }
}
foreach ($t in $ExcludeTags) {
  if ($t -and $t.Trim() -ne "") {
    $cmd = "$cmd --exclude-tag `"$($t.Trim())`""
  }
}
Write-Host "Running engine: $cmd"

Invoke-Expression $cmd
