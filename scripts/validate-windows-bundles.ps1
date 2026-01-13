<#
.SYNOPSIS
  Validate Windows desktop installer bundle outputs are present and (when signing is configured) Authenticode-signed.

.DESCRIPTION
  In release CI, we want to fail early if the Windows desktop installers were not
  produced (publishing an empty Windows release), and to ensure that when
  signing is configured the produced installers are Authenticode-signed.

  By default this script searches common Tauri output locations:
    - apps/desktop/src-tauri/target/**/release/bundle/nsis/*.exe
    - apps/desktop/src-tauri/target/**/release/bundle/msi/*.msi
    - target/**/release/bundle/nsis/*.exe
    - target/**/release/bundle/msi/*.msi

  You can override discovery by providing -ExePath and/or -MsiPath.

.PARAMETER ExePath
  Optional path(s) to NSIS installer .exe files, directories containing them,
  or wildcard patterns. When provided, overrides default NSIS discovery.

.PARAMETER MsiPath
  Optional path(s) to MSI installer .msi files, directories containing them,
  or wildcard patterns. When provided, overrides default MSI discovery.

.EXAMPLE
  pwsh ./scripts/validate-windows-bundles.ps1

.EXAMPLE
  pwsh ./scripts/validate-windows-bundles.ps1 -ExePath apps/desktop/src-tauri/target/release/bundle/nsis/*.exe

.NOTES
  If the environment variable WINDOWS_CERTIFICATE is set (non-empty), this
  script will run:
    signtool verify /pa /all <installer>
  against every discovered installer and will fail if any are unsigned or have
  invalid signatures.
#>

[CmdletBinding()]
param(
  [string[]]$ExePath = @(),
  [string[]]$MsiPath = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Get-RepoRoot {
  # scripts/validate-windows-bundles.ps1 lives in <repoRoot>/scripts/
  $root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")
  return $root.Path
}

function Expand-FileInputs {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Inputs,

    [Parameter(Mandatory = $true)]
    [ValidateSet(".exe", ".msi")]
    [string]$Extension,

    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $out = New-Object System.Collections.Generic.List[System.IO.FileInfo]

  foreach ($raw in $Inputs) {
    if ([string]::IsNullOrWhiteSpace($raw)) { continue }

    $p = [Environment]::ExpandEnvironmentVariables($raw.Trim())

    # Treat relative paths/patterns as relative to the repository root (not the
    # caller's cwd, which can vary across CI steps).
    $candidatePath = $p
    if (-not [System.IO.Path]::IsPathRooted($candidatePath)) {
      $candidatePath = Join-Path $RepoRoot $candidatePath
    }

    if (Test-Path -LiteralPath $candidatePath) {
      $item = Get-Item -LiteralPath $candidatePath
      if ($item.PSIsContainer) {
        $files = Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Filter "*$Extension" -ErrorAction SilentlyContinue
        foreach ($f in $files) { $out.Add($f) }
      } else {
        if ($item.Extension -ieq $Extension) {
          $out.Add([System.IO.FileInfo]$item)
        } else {
          throw "Path does not have expected extension '$Extension': $candidatePath"
        }
      }
      continue
    }

    # Fall back to wildcard expansion (e.g. apps/**/nsis/*.exe).
    $matches = @(Get-ChildItem -Path $candidatePath -File -ErrorAction SilentlyContinue)
    $matches = @($matches | Where-Object { $_.Extension -ieq $Extension })

    if ($matches.Count -eq 0) {
      throw "No files matched '$raw' (resolved as '$candidatePath')"
    }

    foreach ($m in $matches) { $out.Add($m) }
  }

  return @($out | Sort-Object FullName -Unique)
}

function Find-BundleFiles {
  param(
    [Parameter(Mandatory = $true)]
    [string]$TargetRoot,

    [Parameter(Mandatory = $true)]
    [ValidateSet("nsis", "nsis-web", "msi")]
    [string]$BundleKind,

    [Parameter(Mandatory = $true)]
    [ValidateSet(".exe", ".msi")]
    [string]$Extension
  )

  if (-not (Test-Path -LiteralPath $TargetRoot)) {
    return @()
  }

  # We intentionally search for bundle directories first to avoid walking the
  # entire Cargo target tree for all files.
  $bundleDirs = @(Get-ChildItem -LiteralPath $TargetRoot -Recurse -Directory -Filter $BundleKind -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -match "[\\\\/](release)[\\\\/](bundle)[\\\\/]$BundleKind$" })

  $out = New-Object System.Collections.Generic.List[System.IO.FileInfo]
  foreach ($dir in $bundleDirs) {
    $files = Get-ChildItem -LiteralPath $dir.FullName -Recurse -File -Filter "*$Extension" -ErrorAction SilentlyContinue
    foreach ($f in $files) { $out.Add($f) }
  }

  return @($out | Sort-Object FullName -Unique)
}

function Get-SignToolPath {
  $cmd = Get-Command signtool -ErrorAction SilentlyContinue
  if ($null -ne $cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source)) {
    return $cmd.Source
  }

  $pf86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
  if ([string]::IsNullOrWhiteSpace($pf86)) {
    $pf86 = $env:ProgramFiles
  }
  if ([string]::IsNullOrWhiteSpace($pf86)) {
    return $null
  }

  $kits10 = [System.IO.Path]::Combine($pf86, "Windows Kits", "10", "bin")
  $kits81 = [System.IO.Path]::Combine($pf86, "Windows Kits", "8.1", "bin")

  foreach ($kits in @($kits10, $kits81)) {
    if (-not (Test-Path -LiteralPath $kits)) { continue }

    # Some SDK installs expose signtool directly under bin\\x64.
    foreach ($arch in @("x64", "x86", "arm64")) {
      $direct = [System.IO.Path]::Combine($kits, $arch, "signtool.exe")
      if (Test-Path -LiteralPath $direct) {
        return $direct
      }
    }

    # Typical layout: bin\\<version>\\x64\\signtool.exe (prefer newest version, x64).
    $versionDirs = @(Get-ChildItem -LiteralPath $kits -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    foreach ($v in $versionDirs) {
      foreach ($arch in @("x64", "x86", "arm64")) {
        $p = [System.IO.Path]::Combine($v.FullName, $arch, "signtool.exe")
        if (Test-Path -LiteralPath $p) {
          return $p
        }
      }
    }
  }

  return $null
}

function Assert-Signed {
  param(
    [Parameter(Mandatory = $true)]
    [string]$SignToolPath,

    [Parameter(Mandatory = $true)]
    [System.IO.FileInfo]$File
  )

  Write-Host "signtool verify: $($File.FullName)"
  & $SignToolPath verify /pa /all $File.FullName
  if ($LASTEXITCODE -ne 0) {
    throw "Authenticode signature verification failed for: $($File.FullName)"
  }
}

$repoRoot = Get-RepoRoot

Push-Location $repoRoot
try {
  $searchRoots = New-Object System.Collections.Generic.List[string]
  # Prefer CARGO_TARGET_DIR when set (some CI environments export it), but always
  # include the default workspace roots.
  if (-not [string]::IsNullOrWhiteSpace($env:CARGO_TARGET_DIR)) {
    $cargoTarget = $env:CARGO_TARGET_DIR
    if (-not [System.IO.Path]::IsPathRooted($cargoTarget)) {
      $cargoTarget = Join-Path $repoRoot $cargoTarget
    }
    $searchRoots.Add($cargoTarget)
  }
  $searchRoots.Add([System.IO.Path]::Combine($repoRoot, "apps", "desktop", "src-tauri", "target"))
  $searchRoots.Add([System.IO.Path]::Combine($repoRoot, "apps", "desktop", "target"))
  $searchRoots.Add([System.IO.Path]::Combine($repoRoot, "target"))
  $searchRoots = @($searchRoots | Where-Object { Test-Path -LiteralPath $_ } | Sort-Object -Unique)

  $exeInstallers = @()
  $msiInstallers = @()

  if ($ExePath.Count -gt 0) {
    $exeInstallers = Expand-FileInputs -Inputs $ExePath -Extension ".exe" -RepoRoot $repoRoot
  } else {
    foreach ($root in $searchRoots) {
      $exeInstallers += Find-BundleFiles -TargetRoot $root -BundleKind "nsis" -Extension ".exe"
      $exeInstallers += Find-BundleFiles -TargetRoot $root -BundleKind "nsis-web" -Extension ".exe"
    }
    $exeInstallers = @($exeInstallers | Sort-Object FullName -Unique)
  }

  if ($MsiPath.Count -gt 0) {
    $msiInstallers = Expand-FileInputs -Inputs $MsiPath -Extension ".msi" -RepoRoot $repoRoot
  } else {
    foreach ($root in $searchRoots) {
      $msiInstallers += Find-BundleFiles -TargetRoot $root -BundleKind "msi" -Extension ".msi"
    }
    $msiInstallers = @($msiInstallers | Sort-Object FullName -Unique)
  }

  # Exclude embedded WebView2 helper installers; we only care about the Formula installers.
  $exeInstallers = @(
    $exeInstallers | Where-Object { $_.Name -notmatch '^(?i)MicrosoftEdgeWebview2' }
  )

  $totalInstallers = $exeInstallers.Count + $msiInstallers.Count

  Write-Host "Windows bundle validation"
  Write-Host "RepoRoot: $repoRoot"
  if ($searchRoots.Count -gt 0) {
    Write-Host ("SearchRoots: {0}" -f ($searchRoots -join ", "))
  }
  Write-Host ""
  Write-Host ("NSIS installers (.exe): {0}" -f $exeInstallers.Count)
  foreach ($f in $exeInstallers) { Write-Host ("  - {0}" -f $f.FullName) }
  Write-Host ("MSI installers (.msi):  {0}" -f $msiInstallers.Count)
  foreach ($f in $msiInstallers) { Write-Host ("  - {0}" -f $f.FullName) }
  Write-Host ""

  if ($totalInstallers -eq 0) {
    throw "No Windows installer artifacts were found (.exe and .msi are both missing). Ensure the release build produces installers under release/bundle/(nsis|msi)."
  }

  $signingConfigured = -not [string]::IsNullOrWhiteSpace($env:WINDOWS_CERTIFICATE)
  if ($signingConfigured) {
    Write-Host "Signing configuration detected (WINDOWS_CERTIFICATE is set). Verifying Authenticode signatures..."
    $signtoolPath = Get-SignToolPath
    if ([string]::IsNullOrWhiteSpace($signtoolPath)) {
      throw "WINDOWS_CERTIFICATE is set but signtool.exe was not found in PATH or Windows SDK locations. Install the Windows SDK or add signtool to PATH."
    }

    foreach ($installer in @($exeInstallers + $msiInstallers)) {
      Assert-Signed -SignToolPath $signtoolPath -File $installer
    }
  } else {
    Write-Host "Signing not configured (WINDOWS_CERTIFICATE is empty). Skipping signature verification."
  }

  Write-Host "Windows bundle validation succeeded."
} finally {
  Pop-Location
}
