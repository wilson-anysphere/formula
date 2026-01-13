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

.PARAMETER BundleDir
  Optional path to a Tauri bundle directory (…/release/bundle). When provided,
  the script will only search within this bundle directory for:
    - msi/**/*.msi
    - nsis/**/*.exe
    - nsis-web/**/*.exe
  This is useful in CI matrix jobs where the expected output directory is known.

.PARAMETER RequireExe
  Require at least one NSIS installer (.exe). By default, the script only
  requires at least one installer of either type.

.PARAMETER RequireMsi
  Require at least one MSI installer (.msi). By default, the script only
  requires at least one installer of either type.

.EXAMPLE
  pwsh ./scripts/validate-windows-bundles.ps1

.EXAMPLE
  pwsh ./scripts/validate-windows-bundles.ps1 -ExePath apps/desktop/src-tauri/target/release/bundle/nsis/*.exe

.NOTES
  If the environment variable WINDOWS_CERTIFICATE is set (non-empty), this
  script will run:
    signtool verify /pa /all /v <installer>
  against every discovered installer and will fail if any are unsigned or have
  invalid signatures. It also enforces that signatures are timestamped so they
  remain valid after the signing certificate expires.
#>

[CmdletBinding()]
param(
  [string[]]$ExePath = @(),
  [string[]]$MsiPath = @(),
  [string]$BundleDir = "",
  [switch]$RequireExe,
  [switch]$RequireMsi
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

  Write-Host "::group::signtool verify /pa /all $($File.Name)"
  $output = & $SignToolPath verify /pa /all /v $File.FullName 2>&1 | Out-String
  $exitCode = $LASTEXITCODE
  if (-not [string]::IsNullOrWhiteSpace($output)) {
    Write-Host $output.TrimEnd()
  }
  Write-Host "::endgroup::"

  if ($exitCode -ne 0) {
    throw "Authenticode signature verification failed for: $($File.FullName) (signtool exit code $exitCode)"
  }

  # Ensure the signature is timestamped so it remains valid after the signing
  # certificate expires.
  $outLc = $output.ToLowerInvariant()
  if ($outLc.Contains("not timestamped")) {
    throw "Authenticode signature is not timestamped for: $($File.FullName)"
  }
  if (-not ($outLc.Contains("signature is timestamped") -or $outLc.Contains("the signature is timestamped") -or $outLc.Contains("timestamp verified by"))) {
    throw "Unable to determine Authenticode timestamp status for: $($File.FullName) (expected signtool output to mention a timestamp)"
  }
}

$repoRoot = Get-RepoRoot

Push-Location $repoRoot
try {
  $exeInstallers = @()
  $msiInstallers = @()

  $resolvedBundleDir = ""
  $searchRoots = @()

  if (-not [string]::IsNullOrWhiteSpace($BundleDir)) {
    if ($ExePath.Count -gt 0 -or $MsiPath.Count -gt 0) {
      throw "Use either -BundleDir or -ExePath/-MsiPath overrides, not both."
    }

    $candidate = [Environment]::ExpandEnvironmentVariables($BundleDir.Trim())
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
      $candidate = Join-Path $repoRoot $candidate
    }
    if (-not (Test-Path -LiteralPath $candidate)) {
      throw "BundleDir not found: $candidate"
    }
    $bundleItem = Get-Item -LiteralPath $candidate
    if (-not $bundleItem.PSIsContainer) {
      throw "BundleDir must be a directory: $candidate"
    }
    $resolvedBundleDir = $bundleItem.FullName

    $exeInstallers += @(Get-ChildItem -LiteralPath (Join-Path $resolvedBundleDir "nsis") -Recurse -File -Filter "*.exe" -ErrorAction SilentlyContinue)
    $exeInstallers += @(Get-ChildItem -LiteralPath (Join-Path $resolvedBundleDir "nsis-web") -Recurse -File -Filter "*.exe" -ErrorAction SilentlyContinue)
    $msiInstallers += @(Get-ChildItem -LiteralPath (Join-Path $resolvedBundleDir "msi") -Recurse -File -Filter "*.msi" -ErrorAction SilentlyContinue)

    $exeInstallers = @($exeInstallers | Sort-Object FullName -Unique)
    $msiInstallers = @($msiInstallers | Sort-Object FullName -Unique)
  } else {
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
  }

  # Exclude embedded WebView2 helper installers; we only care about the Formula installers.
  $exeInstallers = @(
    $exeInstallers | Where-Object { $_.Name -notmatch '^(?i)MicrosoftEdgeWebview2' }
  )

  $totalInstallers = $exeInstallers.Count + $msiInstallers.Count

  Write-Host "Windows bundle validation"
  Write-Host "RepoRoot: $repoRoot"
  if (-not [string]::IsNullOrWhiteSpace($resolvedBundleDir)) {
    Write-Host "BundleDir: $resolvedBundleDir"
  } elseif ($searchRoots.Count -gt 0) {
    Write-Host ("SearchRoots: {0}" -f ($searchRoots -join ", "))
  }
  Write-Host ""
  Write-Host ("NSIS installers (.exe): {0}" -f $exeInstallers.Count)
  foreach ($f in $exeInstallers) { Write-Host ("  - {0}" -f $f.FullName) }
  Write-Host ("MSI installers (.msi):  {0}" -f $msiInstallers.Count)
  foreach ($f in $msiInstallers) { Write-Host ("  - {0}" -f $f.FullName) }
  Write-Host ""

  $searchHint = ""
  if (-not [string]::IsNullOrWhiteSpace($resolvedBundleDir)) {
    $searchHint = $resolvedBundleDir
  } elseif ($searchRoots.Count -gt 0) {
    $searchHint = ($searchRoots -join ", ")
  } else {
    $searchHint = "<none>"
  }

  if ($RequireExe -and $exeInstallers.Count -eq 0) {
    throw "Missing required Windows installer artifact type: NSIS .exe. Common causes:`n- NSIS missing (makensis.exe not on PATH) so EXE bundling was skipped/failed`n- Tauri NSIS bundler does not support this target`nSearched: $searchHint"
  }
  if ($RequireMsi -and $msiInstallers.Count -eq 0) {
    throw "Missing required Windows installer artifact type: MSI .msi. Common causes:`n- WiX Toolset missing (candle.exe/light.exe not on PATH) so MSI bundling was skipped/failed`n- Tauri MSI bundler does not support this target`nSearched: $searchHint"
  }
  if (-not $RequireExe -and -not $RequireMsi -and $totalInstallers -eq 0) {
    throw "No Windows installer artifacts were found (.exe and .msi are both missing). Ensure the release build produces installers under release/bundle/(nsis|nsis-web|msi)."
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
