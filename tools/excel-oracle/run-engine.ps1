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

.NOTES
  Until a real engine CLI is wired up, this script will fail fast with a clear
  error message.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$CasesPath,

  [Parameter(Mandatory = $true)]
  [string]$OutPath,

  [string]$EngineCommand
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $EngineCommand) {
  $EngineCommand = $env:FORMULA_ENGINE_CMD
}

if (-not $EngineCommand) {
  throw "No engine command provided. Set -EngineCommand or env var FORMULA_ENGINE_CMD to a CLI that can evaluate cases.json and write results JSON."
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
Write-Host "Running engine: $cmd"

Invoke-Expression $cmd

