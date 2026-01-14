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

# `RUSTUP_TOOLCHAIN` overrides the repo's `rust-toolchain.toml` pin. Some environments set it
# globally (often to `stable`), which would bypass the pinned toolchain and reintroduce drift for
# this repo when running `cargo` directly.
if ($env:RUSTUP_TOOLCHAIN -and (Test-Path -LiteralPath (Join-Path $repoRoot "rust-toolchain.toml"))) {
  Remove-Item Env:RUSTUP_TOOLCHAIN -ErrorAction SilentlyContinue
}

$defaultGlobalCargoHome = Join-Path ([Environment]::GetFolderPath("UserProfile")) ".cargo"
$cargoHomeNorm = if ($env:CARGO_HOME) { $env:CARGO_HOME.TrimEnd('\', '/') } else { "" }
$defaultGlobalCargoHomeNorm = $defaultGlobalCargoHome.TrimEnd('\', '/')
if (
  [string]::IsNullOrWhiteSpace($env:CARGO_HOME) -or
  (
    -not $env:CI -and
    -not $env:FORMULA_ALLOW_GLOBAL_CARGO_HOME -and
    $cargoHomeNorm -eq $defaultGlobalCargoHomeNorm
  )
) {
  $env:CARGO_HOME = Join-Path $repoRoot "target/cargo-home"
}

New-Item -ItemType Directory -Force -Path $env:CARGO_HOME | Out-Null

# Ensure tools installed via `cargo install` under this CARGO_HOME are available.
$cargoBinDir = Join-Path $env:CARGO_HOME "bin"
New-Item -ItemType Directory -Force -Path $cargoBinDir | Out-Null
$pathEntries = $env:Path -split ';'
if (-not ($pathEntries -contains $cargoBinDir)) {
  $env:Path = "$cargoBinDir;$env:Path"
}

# Concurrency defaults: keep Rust builds stable on high-core-count multi-agent hosts.
#
# Prefer explicit overrides, but default to a conservative job count when unset. On very
# high core-count hosts, linking (lld) can spawn many threads per link step; combining that
# with Cargo-level parallelism can exceed sandbox process/thread limits and cause flaky
# "Resource temporarily unavailable" failures.
$cpuCount = [Environment]::ProcessorCount
$defaultJobsInt = if ($cpuCount -ge 64) { 2 } else { 4 }
$jobsRaw = if ($env:FORMULA_CARGO_JOBS) { $env:FORMULA_CARGO_JOBS } elseif ($env:CARGO_BUILD_JOBS) { $env:CARGO_BUILD_JOBS } else { $defaultJobsInt.ToString() }
$jobsInt = 0
if (-not [int]::TryParse($jobsRaw, [ref]$jobsInt) -or $jobsInt -lt 1) { $jobsInt = $defaultJobsInt }
$jobs = $jobsInt.ToString()

$env:CARGO_BUILD_JOBS = $jobs
if (-not $env:MAKEFLAGS) { $env:MAKEFLAGS = "-j$jobs" }
if (-not $env:CARGO_PROFILE_DEV_CODEGEN_UNITS) { $env:CARGO_PROFILE_DEV_CODEGEN_UNITS = $jobs }
if (-not $env:CARGO_PROFILE_TEST_CODEGEN_UNITS) { $env:CARGO_PROFILE_TEST_CODEGEN_UNITS = $jobs }
if (-not $env:CARGO_PROFILE_RELEASE_CODEGEN_UNITS) { $env:CARGO_PROFILE_RELEASE_CODEGEN_UNITS = $jobs }
if (-not $env:CARGO_PROFILE_BENCH_CODEGEN_UNITS) { $env:CARGO_PROFILE_BENCH_CODEGEN_UNITS = $jobs }
if (-not $env:RAYON_NUM_THREADS) {
  $env:RAYON_NUM_THREADS = if ($env:FORMULA_RAYON_NUM_THREADS) { $env:FORMULA_RAYON_NUM_THREADS } else { $jobs }
}

# Some environments configure Cargo globally with `build.rustc-wrapper`. When the wrapper is
# unavailable/misconfigured, builds can fail even for `cargo metadata`. Default to disabling any
# configured wrapper unless the user explicitly overrides it.
$rustcWrapper = [Environment]::GetEnvironmentVariable("RUSTC_WRAPPER")
if ($null -eq $rustcWrapper) { $rustcWrapper = [Environment]::GetEnvironmentVariable("CARGO_BUILD_RUSTC_WRAPPER") }
if ($null -eq $rustcWrapper) { $rustcWrapper = "" }

$rustcWorkspaceWrapper = [Environment]::GetEnvironmentVariable("RUSTC_WORKSPACE_WRAPPER")
if ($null -eq $rustcWorkspaceWrapper) {
  $rustcWorkspaceWrapper = [Environment]::GetEnvironmentVariable("CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER")
}
if ($null -eq $rustcWorkspaceWrapper) { $rustcWorkspaceWrapper = "" }

$env:RUSTC_WRAPPER = $rustcWrapper
$env:RUSTC_WORKSPACE_WRAPPER = $rustcWorkspaceWrapper
# Cargo can also read wrapper config via `CARGO_BUILD_RUSTC_WRAPPER`. Set it explicitly so a global
# Cargo config cannot unexpectedly re-enable a flaky wrapper when the user didn't opt in.
$env:CARGO_BUILD_RUSTC_WRAPPER = $rustcWrapper
$env:CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER = $rustcWorkspaceWrapper

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
