<#
.SYNOPSIS
  Validate Windows desktop installer bundle outputs are present, contain required desktop integration metadata, and (when signing is configured) are Authenticode-signed.

.DESCRIPTION
  In release CI, we want to fail early if the Windows desktop installers were not
  produced (publishing an empty Windows release), and to ensure that when
  signing is configured the produced installers are Authenticode-signed.
 
  We also validate that the built installers include registry entries for:
    - file associations configured in `apps/desktop/src-tauri/tauri.conf.json` bundle.fileAssociations
      (e.g. `.xlsx`, `.csv`, `.parquet`, ...)
    - the `formula://` URL protocol handler.

  By default this script searches common Tauri output locations (including
  workspace target roots and per-target-triple subdirectories):
    - apps/desktop/src-tauri/target/release/bundle/(nsis|nsis-web|msi)/*
    - apps/desktop/src-tauri/target/*/release/bundle/(nsis|nsis-web|msi)/*
    - apps/desktop/target/release/bundle/(nsis|nsis-web|msi)/*
    - apps/desktop/target/*/release/bundle/(nsis|nsis-web|msi)/*
    - target/release/bundle/(nsis|nsis-web|msi)/*
    - target/*/release/bundle/(nsis|nsis-web|msi)/*
    - $env:CARGO_TARGET_DIR/release/bundle/(nsis|nsis-web|msi)/* (when set)
    - $env:CARGO_TARGET_DIR/*/release/bundle/(nsis|nsis-web|msi)/* (when set)

  You can override discovery by providing -ExePath and/or -MsiPath.

.PARAMETER ExePath
  Optional path(s) to NSIS installer .exe files, directories containing them,
  or wildcard patterns. Directories are searched recursively. When provided,
  overrides default NSIS discovery.

.PARAMETER MsiPath
  Optional path(s) to MSI installer .msi files, directories containing them,
  or wildcard patterns. Directories are searched recursively. When provided,
  overrides default MSI discovery.

.PARAMETER BundleDir
  Optional path to a Tauri bundle directory (.../release/bundle). When provided,
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
  pwsh -NoProfile -ExecutionPolicy Bypass -File ./scripts/validate-windows-bundles.ps1

.EXAMPLE
  pwsh -NoProfile -ExecutionPolicy Bypass -File ./scripts/validate-windows-bundles.ps1 -ExePath apps/desktop/src-tauri/target/release/bundle/nsis/*.exe

.EXAMPLE
  pwsh -NoProfile -ExecutionPolicy Bypass -File ./scripts/validate-windows-bundles.ps1 -BundleDir apps/desktop/src-tauri/target/x86_64-pc-windows-msvc/release/bundle -RequireExe -RequireMsi

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

$runningOnWindows = $false
try {
  $runningOnWindows = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
    [System.Runtime.InteropServices.OSPlatform]::Windows
  )
} catch {
  # Best-effort fallback for older PowerShell/.NET environments.
  $runningOnWindows = ($env:OS -eq "Windows_NT")
}

if (-not $runningOnWindows) {
  throw "scripts/validate-windows-bundles.ps1 must be run on Windows (requires signtool.exe and Windows Installer COM APIs)."
}

function Get-RepoRoot {
  # scripts/validate-windows-bundles.ps1 lives in <repoRoot>/scripts/
  $root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")
  return $root.Path
}

function Get-TauriConfPath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $raw = [string]$env:FORMULA_TAURI_CONF_PATH
  if (-not [string]::IsNullOrWhiteSpace($raw)) {
    $p = $raw.Trim()
    if ([System.IO.Path]::IsPathRooted($p)) {
      return $p
    }
    return (Join-Path $RepoRoot $p)
  }

  return (Join-Path $RepoRoot "apps/desktop/src-tauri/tauri.conf.json")
}

function Get-ExpectedTauriVersion {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $tauriConfPath = Get-TauriConfPath -RepoRoot $RepoRoot
  if (-not (Test-Path -LiteralPath $tauriConfPath)) {
    throw "Missing Tauri config: $tauriConfPath"
  }

  $conf = Get-Content -Raw -LiteralPath $tauriConfPath | ConvertFrom-Json
  $v = [string]$conf.version
  if ([string]::IsNullOrWhiteSpace($v)) {
    throw "Expected $tauriConfPath to contain a non-empty `"version`" field."
  }
  return $v.Trim()
}

function Normalize-Guid {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Value
  )
 
  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ""
  }
 
  $s = $Value.Trim()
  # MSI tables sometimes store GUIDs wrapped in braces.
  $s = $s.TrimStart("{").TrimEnd("}")
 
  try {
    return ([Guid]$s).ToString("D").ToLowerInvariant()
  } catch {
    return ""
  }
}
 
function Get-ExpectedWixUpgradeCode {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $tauriConfPath = Get-TauriConfPath -RepoRoot $RepoRoot
  if (-not (Test-Path -LiteralPath $tauriConfPath)) {
    return ""
  }
 
  try {
    $conf = Get-Content -Raw -LiteralPath $tauriConfPath | ConvertFrom-Json
  } catch {
    return ""
  }
 
  $bundleProp = $conf.PSObject.Properties["bundle"]
  if ($null -eq $bundleProp -or $null -eq $bundleProp.Value) { return "" }
  $bundle = $bundleProp.Value
 
  $windowsProp = $bundle.PSObject.Properties["windows"]
  if ($null -eq $windowsProp -or $null -eq $windowsProp.Value) { return "" }
  $windows = $windowsProp.Value
 
  $wixProp = $windows.PSObject.Properties["wix"]
  if ($null -eq $wixProp -or $null -eq $wixProp.Value) { return "" }
  $wix = $wixProp.Value
 
  $upgradeProp = $wix.PSObject.Properties["upgradeCode"]
  if ($null -eq $upgradeProp -or $null -eq $upgradeProp.Value) { return "" }
  $v = [string]$upgradeProp.Value
  if ([string]::IsNullOrWhiteSpace($v)) { return "" }
 
  return $v.Trim()
}

function Get-ExpectedProductName {
  param(
    [Parameter(Mandatory = $true)]
    [string]$RepoRoot
  )

  $tauriConfPath = Get-TauriConfPath -RepoRoot $RepoRoot
  if (-not (Test-Path -LiteralPath $tauriConfPath)) {
    return ""
  }
 
  try {
    $conf = Get-Content -Raw -LiteralPath $tauriConfPath | ConvertFrom-Json
  } catch {
    return ""
  }
 
  $v = [string]$conf.productName
  if ([string]::IsNullOrWhiteSpace($v)) {
    return ""
  }
  return $v.Trim()
}

function Normalize-Version {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Version
  )

  if ([string]::IsNullOrWhiteSpace($Version)) {
    return ""
  }

  # Extract the first numeric dotted version prefix (e.g. 1.2.3 or 1.2.3.4) from
  # strings like "1.2.3.0 (some text)".
  $m = [regex]::Match($Version, "\d+(?:\.\d+){1,3}")
  if (-not $m.Success) {
    return ""
  }

  $parts = $m.Value.Split(".")
  # Windows file versions sometimes include a trailing ".0" (4-part version).
  while ($parts.Length -gt 3 -and $parts[$parts.Length - 1] -eq "0") {
    $parts = $parts[0..($parts.Length - 2)]
  }
  return ($parts -join ".")
}

function Assert-VersionMatch {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ArtifactPath,
    [Parameter(Mandatory = $true)]
    [string]$FoundVersion,
    [Parameter(Mandatory = $true)]
    [string]$ExpectedVersion,
    [Parameter(Mandatory = $true)]
    [string]$Context
  )

  $expectedNorm = Normalize-Version -Version $ExpectedVersion
  $foundNorm = Normalize-Version -Version $FoundVersion

  if ([string]::IsNullOrWhiteSpace($expectedNorm) -or [string]::IsNullOrWhiteSpace($foundNorm)) {
    throw "Unable to parse version for $Context.`n- Artifact: $ArtifactPath`n- Expected: $ExpectedVersion`n- Found: $FoundVersion"
  }

  if ($expectedNorm -ne $foundNorm) {
    throw "Windows bundle version mismatch detected ($Context).`n- Artifact: $ArtifactPath`n- Expected (tauri.conf.json version): $ExpectedVersion`n- Found: $FoundVersion"
  }
}

function Get-MsiProperty {
  param(
    [Parameter(Mandatory = $true)]
    [string]$MsiPath,
    [Parameter(Mandatory = $true)]
    [string]$PropertyName
  )

  try {
    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = $installer.OpenDatabase($MsiPath, 0)
    $query = "SELECT `Value` FROM `Property` WHERE `Property`='$PropertyName'"
    $view = $database.OpenView($query)
    $view.Execute()
    $record = $view.Fetch()
    if ($null -eq $record) {
      return $null
    }
    return $record.StringData(1)
  } catch {
    return $null
  }
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

function Find-BundleKindDirsFallback {
  <#
    Best-effort fallback discovery for bundle directories when the expected Cargo/Tauri layout
    is not present.

    Historical implementations used `Get-ChildItem -Recurse` over the entire Cargo target
    directory. Target trees can be extremely large (build scripts, deps, incremental, etc),
    and a full recursive enumeration can add minutes of overhead in CI once builds have run.

    This helper performs a bounded directory walk and prunes well-known large directories.
  #>
  param(
    [Parameter(Mandatory = $true)]
    [string]$TargetRoot,

    [Parameter(Mandatory = $true)]
    [string]$BundleKind,

    [int]$MaxDepth = 8
  )

  if (-not (Test-Path -LiteralPath $TargetRoot -PathType Container)) {
    return @()
  }

  # Common large directories inside Cargo targets (and other repo trees).
  $skipNames = @(
    "build",
    "deps",
    "incremental",
    ".fingerprint",
    "debug",
    ".git",
    "node_modules",
    "dist",
    "build",
    "coverage",
    "target",
    ".pnpm-store",
    ".turbo",
    ".cache",
    ".vite",
    "security-report",
    "test-results",
    "playwright-report"
  )

  $found = New-Object System.Collections.Generic.List[string]

  function Walk-Dir {
    param(
      [Parameter(Mandatory = $true)]
      [string]$Path,

      [Parameter(Mandatory = $true)]
      [int]$Depth
    )

    if ($Depth -gt $MaxDepth) {
      return
    }

    $dirs = @()
    try {
      $dirs = @(Get-ChildItem -LiteralPath $Path -Directory -ErrorAction SilentlyContinue)
    } catch {
      $dirs = @()
    }

    foreach ($d in $dirs) {
      # Avoid junction/symlink loops.
      try {
        if (($d.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
      } catch {
        # Best-effort; if attribute checks fail, continue.
      }

      if ($skipNames -contains $d.Name) { continue }

      if ($d.Name -ieq $BundleKind -and $d.FullName -match "[\\\\/](release)[\\\\/](bundle)[\\\\/]$BundleKind$") {
        $found.Add($d.FullName) | Out-Null
        continue
      }

      Walk-Dir -Path $d.FullName -Depth ($Depth + 1)
    }
  }

  Walk-Dir -Path $TargetRoot -Depth 0

  return @($found | Sort-Object -Unique)
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

  # Prefer checking the expected Tauri/Cargo output locations:
  #   - <target>/release/bundle/<kind>
  #   - <target>/<triple>/release/bundle/<kind>
  #
  # Avoid recursive enumeration of the entire Cargo target tree, which can be very large in CI.
  $bundleDirs = New-Object System.Collections.Generic.List[string]

  $native = Join-Path $TargetRoot (Join-Path "release" (Join-Path "bundle" $BundleKind))
  if (Test-Path -LiteralPath $native -PathType Container) {
    $bundleDirs.Add($native)
  }

  $targetEntries = @()
  try {
    $targetEntries = @(Get-ChildItem -LiteralPath $TargetRoot -Directory -ErrorAction SilentlyContinue)
  } catch {
    $targetEntries = @()
  }
  foreach ($ent in $targetEntries) {
    $maybe = Join-Path $ent.FullName (Join-Path "release" (Join-Path "bundle" $BundleKind))
    if (Test-Path -LiteralPath $maybe -PathType Container) {
      $bundleDirs.Add($maybe)
    }
  }

  $out = New-Object System.Collections.Generic.List[System.IO.FileInfo]
  foreach ($dirPath in $bundleDirs) {
    $files = Get-ChildItem -LiteralPath $dirPath -Recurse -File -Filter "*$Extension" -ErrorAction SilentlyContinue
    foreach ($f in $files) { $out.Add($f) }
  }

  if ($out.Count -gt 0) {
    return @($out | Sort-Object FullName -Unique)
  }

  # Fall back to a recursive search for bundle directories for any non-standard layouts.
  $bundleDirsFallback = Find-BundleKindDirsFallback -TargetRoot $TargetRoot -BundleKind $BundleKind -MaxDepth 8
  foreach ($dirPath in $bundleDirsFallback) {
    $files = Get-ChildItem -LiteralPath $dirPath -Recurse -File -Filter "*$Extension" -ErrorAction SilentlyContinue
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

function Get-SevenZipPath {
  $cmd = Get-Command 7z.exe -ErrorAction SilentlyContinue
  if ($null -ne $cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source)) {
    return $cmd.Source
  }
  $cmd = Get-Command 7z -ErrorAction SilentlyContinue
  if ($null -ne $cmd -and -not [string]::IsNullOrWhiteSpace($cmd.Source)) {
    return $cmd.Source
  }

  $candidates = @()
  if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) {
    $candidates += (Join-Path $env:ProgramFiles "7-Zip\\7z.exe")
  }
  if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
    $candidates += (Join-Path ${env:ProgramFiles(x86)} "7-Zip\\7z.exe")
  }
  foreach ($c in $candidates) {
    if (Test-Path -LiteralPath $c) {
      return $c
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

  # Validate desktop integration metadata is present in the produced installers:
  # - file associations (extensions) configured in tauri.conf.json bundle.fileAssociations
  # - the `formula://` URL protocol handler.
  #
  # On Windows, `.xlsx` file associations are typically registered via MSI tables
  # (Extension/ProgId/Verb). This is the most reliable thing to validate in CI.
  #
  # For NSIS `.exe` installers, reliable structured inspection tooling is not always available on
  # GitHub-hosted runners. We do a best-effort streaming marker scan for well-known strings.
  #
  # NOTE: EXE validation is heuristic: it is designed to catch obvious regressions, not to fully
  # prove that the installer will register everything correctly on every machine.
  function Get-ExpectedFileAssociationSpec {
    param(
      [Parameter(Mandatory = $true)]
      [string]$RepoRoot
    )

    $configPath = Get-TauriConfPath -RepoRoot $RepoRoot
    $defaultXlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    $default = [pscustomobject]@{
      Extensions = @("xlsx")
      MimeTypesByExtension = @{ xlsx = $defaultXlsxMime }
      XlsxMimeType = $defaultXlsxMime
    }

    if (-not (Test-Path -LiteralPath $configPath)) {
      return $default
    }

    try {
      $conf = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    } catch {
      Write-Warning "Failed to parse tauri.conf.json for file association expectations: $($_.Exception.Message)"
      return $default
    }

    $bundleProp = $conf.PSObject.Properties["bundle"]
    if ($null -eq $bundleProp -or $null -eq $bundleProp.Value) {
      return $default
    }
    $bundle = $bundleProp.Value

    $assocProp = $bundle.PSObject.Properties["fileAssociations"]
    if ($null -eq $assocProp) {
      return $default
    }
    $fileAssociations = $assocProp.Value
    if ($null -eq $fileAssociations) {
      return $default
    }

    $extSet = New-Object "System.Collections.Generic.HashSet[string]"
    $mimeTypesByExtension = @{}

    foreach ($assoc in @($fileAssociations)) {
      if ($null -eq $assoc) { continue }
      if (-not ($assoc.PSObject.Properties.Name -contains "ext")) { continue }

      # `ext` can be either a string or an array of strings.
      $exts = @($assoc.ext)

      $mime = ""
      if (
        ($assoc.PSObject.Properties.Name -contains "mimeType") -and
        -not [string]::IsNullOrWhiteSpace($assoc.mimeType)
      ) {
        $mime = ($assoc.mimeType).ToString().Trim()
      }

      foreach ($extRaw in $exts) {
        if ($null -eq $extRaw) { continue }
        $ext = $extRaw.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($ext)) { continue }
        $ext = $ext.TrimStart(".").ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($ext)) { continue }

        $null = $extSet.Add($ext)
        if (-not [string]::IsNullOrWhiteSpace($mime) -and -not $mimeTypesByExtension.ContainsKey($ext)) {
          $mimeTypesByExtension[$ext] = $mime
        }
      }
    }

    if ($extSet.Count -eq 0) {
      return $default
    }

    $extensions = @($extSet | Sort-Object -Unique)
    $xlsxMime = $defaultXlsxMime
    if ($mimeTypesByExtension.ContainsKey("xlsx")) {
      $xlsxMime = ($mimeTypesByExtension["xlsx"]).ToString().Trim()
    }

    # Ensure the default has a stable mapping for `.xlsx` even if it was omitted from the config.
    if (-not $mimeTypesByExtension.ContainsKey("xlsx")) {
      $mimeTypesByExtension["xlsx"] = $xlsxMime
    }

    return [pscustomobject]@{
      Extensions = $extensions
      MimeTypesByExtension = $mimeTypesByExtension
      XlsxMimeType = $xlsxMime
    }
  }

  function Get-ExpectedUrlProtocolSpec {
    param(
      [Parameter(Mandatory = $true)]
      [string]$RepoRoot
    )

    $configPath = Get-TauriConfPath -RepoRoot $RepoRoot
    $default = [pscustomobject]@{
      Schemes = @("formula")
    }
 
    if (-not (Test-Path -LiteralPath $configPath)) {
      return $default
    }
 
    try {
      $conf = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    } catch {
      Write-Warning "Failed to parse tauri.conf.json for URL protocol expectations: $($_.Exception.Message)"
      return $default
    }
 
    $pluginsProp = $conf.PSObject.Properties["plugins"]
    if ($null -eq $pluginsProp -or $null -eq $pluginsProp.Value) {
      return $default
    }
    $plugins = $pluginsProp.Value
 
    $deepLinkProp = $plugins.PSObject.Properties["deep-link"]
    if ($null -eq $deepLinkProp -or $null -eq $deepLinkProp.Value) {
      return $default
    }
    $deepLink = $deepLinkProp.Value
 
    $desktopProp = $deepLink.PSObject.Properties["desktop"]
    if ($null -eq $desktopProp -or $null -eq $desktopProp.Value) {
      return $default
    }
    $desktop = $desktopProp.Value
 
    $schemes = @()
 
    # plugins.deep-link.desktop can be either:
    # - a single protocol object with a `schemes` field
    # - an array of protocol objects
    foreach ($proto in @($desktop)) {
      if ($null -eq $proto) { continue }
 
      $schemesProp = $proto.PSObject.Properties["schemes"]
      if ($null -eq $schemesProp -or $null -eq $schemesProp.Value) { continue }
 
      foreach ($s in @($schemesProp.Value)) {
        if ($null -eq $s) { continue }
        $v = $s.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        # Normalize common user input like "formula://" to just the scheme name.
        $v = $v.TrimEnd("/").TrimEnd(":")
        $v = $v.ToLowerInvariant()
        if ($v -match '[:/]') {
          $raw = $s.ToString()
          throw "Invalid deep-link scheme configured in tauri.conf.json. Expected scheme names only (no ':' or '/' characters). Found: $raw"
        }
        if (-not [string]::IsNullOrWhiteSpace($v)) {
          $schemes += $v
        }
      }
    }
    if ($schemes.Count -eq 0) {
      return $default
    }
 
    return [pscustomobject]@{
      Schemes = @($schemes | Sort-Object -Unique)
    }
  }

  function Get-MsiTableNames {
    param(
      [Parameter(Mandatory = $true)]
      [string]$MsiPath
    )
    $installer = New-Object -ComObject WindowsInstaller.Installer
    $db = $installer.OpenDatabase($MsiPath, 0)
    $view = $db.OpenView('SELECT `Name` FROM `_Tables`')
    $view.Execute()
    $names = @()
    while ($true) {
      $rec = $view.Fetch()
      if ($null -eq $rec) { break }
      $names += $rec.StringData(1)
    }
    return $names
  }

  function Get-MsiRows {
    param(
      [Parameter(Mandatory = $true)]
      [string]$MsiPath,
      [Parameter(Mandatory = $true)]
      [string]$Query,
      [Parameter(Mandatory = $true)]
      [int]$ColumnCount
    )

    $installer = New-Object -ComObject WindowsInstaller.Installer
    $db = $installer.OpenDatabase($MsiPath, 0)
    $view = $db.OpenView($Query)
    $view.Execute()

    $rows = @()
    while ($true) {
      $rec = $view.Fetch()
      if ($null -eq $rec) { break }

      $values = @()
      for ($i = 1; $i -le $ColumnCount; $i++) {
        $values += $rec.StringData($i)
      }
      $rows += ,$values
    }
    return $rows
  }

  function Test-MsiRegistryTableForExtension {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Msi,
      [Parameter(Mandatory = $true)]
      [string]$ExtensionNoDot
    )

    $dotExt = "." + $ExtensionNoDot
    $needle = $dotExt
    $needleAlt = $ExtensionNoDot
    $extraNeedles = @()
    # Common ProgIds used when `.xlsx` is associated via registry keys rather than the MSI Extension table.
    if ($ExtensionNoDot -ieq "xlsx") {
      $extraNeedles += @("Excel.Sheet.12")
    }

    try {
      $rows = Get-MsiRows -MsiPath $Msi.FullName -Query 'SELECT `Key`, `Name`, `Value` FROM `Registry`' -ColumnCount 3
    } catch {
      return $false
    }

    foreach ($row in $rows) {
      if ($row.Count -lt 3) { continue }
      foreach ($col in @($row[0], $row[1], $row[2])) {
        if ($null -eq $col) { continue }
        $s = $col.ToString()
        if ($s.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
        if ($s.IndexOf($needleAlt, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
        foreach ($n in $extraNeedles) {
          if ($s.IndexOf($n, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
        }
      }
    }
    return $false
  }

  function Assert-MsiDeclaresFileAssociation {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Msi,
      [Parameter(Mandatory = $true)]
      [string]$ExtensionNoDot,
      [string]$ExpectedMimeType = ""
    )

    Write-Host "File association check (MSI): $($Msi.FullName)"

    $tables = @()
    try {
      $tables = Get-MsiTableNames -MsiPath $Msi.FullName
    } catch {
      throw "Failed to open MSI for inspection: $($Msi.FullName)`n$($_.Exception.Message)"
    }

    $hasRegistryFallback = $tables -contains "Registry"

    if (-not ($tables -contains "Extension")) {
      if ($hasRegistryFallback -and (Test-MsiRegistryTableForExtension -Msi $Msi -ExtensionNoDot $ExtensionNoDot)) {
        Write-Warning "MSI is missing the Extension table, but the Registry table contains strings related to '.$ExtensionNoDot'. Assuming file association metadata is present (best-effort)."
        return
      }
      throw "MSI is missing the Extension table; cannot verify file associations for '.$ExtensionNoDot'. (Check bundle.fileAssociations in tauri.conf.json and the Windows bundler output.)"
    }

    $extRows = Get-MsiRows -MsiPath $Msi.FullName -Query 'SELECT `Extension`, `ProgId_`, `MIME_` FROM `Extension`' -ColumnCount 3
    $foundRow = $null
    foreach ($row in $extRows) {
      if ($row.Count -lt 1) { continue }
      $extVal = if ($null -ne $row[0]) { $row[0] } else { "" }
      $ext = $extVal.Trim().TrimStart(".")
      if ($ext -ieq $ExtensionNoDot) {
        $foundRow = $row
        break
      }
    }

    if ($null -eq $foundRow) {
      if ($hasRegistryFallback -and (Test-MsiRegistryTableForExtension -Msi $Msi -ExtensionNoDot $ExtensionNoDot)) {
        Write-Warning "MSI did not contain an Extension table row for '$ExtensionNoDot', but the Registry table contains strings related to '.$ExtensionNoDot'. Assuming file association metadata is present (best-effort)."
        return
      }
      $present = @(
        $extRows |
          ForEach-Object {
            $v = if ($null -ne $_[0]) { $_[0] } else { "" }
            $v.Trim()
          } |
          Where-Object { $_ -and $_.Trim().Length -gt 0 } |
          Sort-Object -Unique
      )
      $presentText = if ($present.Count -gt 0) { $present -join ", " } else { "(none)" }
      throw "MSI did not declare a file association for '.$ExtensionNoDot'. Expected to find an Extension table row for '$ExtensionNoDot'. Present extensions: $presentText"
    }

    $progIdVal = if ($null -ne $foundRow[1]) { $foundRow[1] } else { "" }
    $progId = $progIdVal.Trim()
    if ([string]::IsNullOrWhiteSpace($progId)) {
      # Some toolchains may register associations via explicit Registry table entries (or other
      # mechanisms) rather than relying on the advertised Extension/ProgId mapping. If the MSI's
      # Registry table contains the extension/progid strings, treat this as best-effort evidence
      # that file associations are still being registered.
      if ($hasRegistryFallback -and (Test-MsiRegistryTableForExtension -Msi $Msi -ExtensionNoDot $ExtensionNoDot)) {
        Write-Warning "MSI Extension table row for '$ExtensionNoDot' exists but ProgId_ is empty. The Registry table contains strings related to '.$ExtensionNoDot'; assuming file association metadata is present (best-effort)."
        return
      }
      throw "MSI Extension table row for '$ExtensionNoDot' exists but ProgId_ is empty. This suggests file association wiring is incomplete."
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedMimeType)) {
      $mimeVal = if ($null -ne $foundRow[2]) { $foundRow[2] } else { "" }
      $mime = $mimeVal.Trim()
      if ([string]::IsNullOrWhiteSpace($mime)) {
        Write-Warning "MSI Extension row for '$ExtensionNoDot' has an empty MIME_ column (expected '$ExpectedMimeType'). File associations may still work via ProgId/registry, but this is unexpected."
      } elseif ($mime -ine $ExpectedMimeType) {
        Write-Warning "MSI Extension row for '$ExtensionNoDot' has MIME_ '$mime' (expected '$ExpectedMimeType')."
      }
    }

    if ($tables -contains "ProgId") {
      $progRows = Get-MsiRows -MsiPath $Msi.FullName -Query 'SELECT `ProgId` FROM `ProgId`' -ColumnCount 1
      $hasProgId = $false
      foreach ($r in $progRows) {
        $v = if ($null -ne $r[0]) { $r[0] } else { "" }
        if ($v.Trim() -ieq $progId) { $hasProgId = $true; break }
      }
      if (-not $hasProgId) {
        Write-Warning "MSI Extension row ProgId_ '$progId' did not appear in the ProgId table. Installer may still work via Registry table entries, but this is unexpected."
      }
    } else {
      Write-Warning "MSI does not include a ProgId table; skipping ProgId validation."
    }

    if ($tables -contains "Verb") {
      # Best-effort: ensure at least one verb is declared for this extension.
      try {
        $verbRows = Get-MsiRows -MsiPath $Msi.FullName -Query 'SELECT `Extension_`, `Verb` FROM `Verb`' -ColumnCount 2
        $hasVerb = $false
        foreach ($r in $verbRows) {
          $v = if ($null -ne $r[0]) { $r[0] } else { "" }
          $ext = $v.Trim().TrimStart(".")
          if ($ext -ieq $ExtensionNoDot) { $hasVerb = $true; break }
        }
        if (-not $hasVerb) {
          Write-Warning "MSI contains a Verb table but no entries for extension '$ExtensionNoDot'. File association may be incomplete."
        }
      } catch {
        Write-Warning "Failed to query MSI Verb table for $($Msi.Name): $($_.Exception.Message)"
      }
    }
  }
 
  function Assert-MsiRegistersUrlProtocol {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Msi,
      [Parameter(Mandatory = $true)]
      [string]$Scheme
    )
 
    $schemeNorm = $Scheme.Trim().TrimEnd("/").TrimEnd(":")
    if ([string]::IsNullOrWhiteSpace($schemeNorm)) {
      throw "URL protocol handler check: expected a non-empty scheme name."
    }
    $schemeNormLc = $schemeNorm.ToLowerInvariant()
 
    Write-Host "URL protocol handler check (MSI): $($Msi.FullName)"
 
    $tables = @()
    try {
      $tables = Get-MsiTableNames -MsiPath $Msi.FullName
    } catch {
      throw "Failed to open MSI for inspection: $($Msi.FullName)`n$($_.Exception.Message)"
    }
 
    if (-not ($tables -contains "Registry")) {
      throw "MSI is missing the Registry table; cannot verify URL protocol handler registration for '$schemeNorm://'."
    }
 
    $rows = Get-MsiRows -MsiPath $Msi.FullName -Query 'SELECT `Root`, `Key`, `Name`, `Value` FROM `Registry`' -ColumnCount 4
 
    $schemeKeyMatch = $false
    $urlProtocolValueMatch = $false
    $schemeValueNames = New-Object System.Collections.Generic.List[string]
    $schemeKeyCandidates = New-Object System.Collections.Generic.List[string]
    foreach ($r in $rows) {
      if ($r.Count -lt 4) { continue }
      $keyRaw = if ($null -ne $r[1]) { $r[1] } else { "" }
      $nameRaw = if ($null -ne $r[2]) { $r[2] } else { "" }

      $key = $keyRaw.Trim().Replace("/", "\").Trim("\")
      $name = $nameRaw.Trim()

      if ([string]::IsNullOrWhiteSpace($key)) { continue }
      if ($key.ToLowerInvariant().Contains($schemeNormLc)) {
        if (-not $schemeKeyCandidates.Contains($key)) {
          $schemeKeyCandidates.Add($key) | Out-Null
        }
      }
 
      # The protocol can be registered either directly under HKCR\<scheme> (Root=0, Key="formula")
      # or under HKCU/HKLM\Software\Classes\<scheme> which merges into HKCR.
      $isSchemeKey =
        ($key -ieq $schemeNorm) -or
        ($key -match "(?i)(^|\\)Software\\Classes\\$([regex]::Escape($schemeNorm))$") -or
        ($key -match "(?i)(^|\\)$([regex]::Escape($schemeNorm))$")
 
      if (-not $isSchemeKey) { continue }
      $schemeKeyMatch = $true

      $prettyName = if ([string]::IsNullOrWhiteSpace($name)) { "(default)" } else { $name }
      if (-not $schemeValueNames.Contains($prettyName)) {
        $schemeValueNames.Add($prettyName) | Out-Null
      }

      if ($name -ieq "URL Protocol") {
        $urlProtocolValueMatch = $true
        break
      }
    }

    if (-not $schemeKeyMatch) {
      $candidatesText = "(none)"
      if ($schemeKeyCandidates.Count -gt 0) {
        $candidatesText = (@($schemeKeyCandidates | Sort-Object -Unique | Select-Object -First 20) -join ", ")
      }
      throw "MSI did not contain any Registry table entries for the '$schemeNorm' protocol key. Expected to see a key like HKCR\\$schemeNorm or (HKCU/HKLM)\\Software\\Classes\\$schemeNorm.`n- Registry keys containing '$schemeNorm': $candidatesText"
    }
    if (-not $urlProtocolValueMatch) {
      $valuesText = "(none)"
      if ($schemeValueNames.Count -gt 0) {
        $valuesText = (@($schemeValueNames | Sort-Object -Unique) -join ", ")
      }
      throw "MSI did not register '$schemeNorm://' as a URL protocol handler. Expected a Registry table value named 'URL Protocol' under HKCR\\$schemeNorm (or equivalent under Software\\Classes\\$schemeNorm).`n- Values found for scheme key: $valuesText"
    }
  }

  function Assert-MsiContainsComplianceArtifacts {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Msi
    )
    Write-Host "Compliance artifact check (MSI): $($Msi.FullName)"

    $tables = @()
    try {
      $tables = Get-MsiTableNames -MsiPath $Msi.FullName
    } catch {
      throw "Failed to open MSI for inspection: $($Msi.FullName)`n$($_.Exception.Message)"
    }

    if (-not ($tables -contains "File")) {
      throw "MSI is missing the File table; cannot validate LICENSE/NOTICE are included."
    }

    $fileRows = Get-MsiRows -MsiPath $Msi.FullName -Query 'SELECT `FileName` FROM `File`' -ColumnCount 1
    $fileNames = @()
    foreach ($row in $fileRows) {
      if ($row.Count -lt 1) { continue }
      $raw = if ($null -ne $row[0]) { $row[0] } else { "" }
      $raw = $raw.Trim()
      if ([string]::IsNullOrWhiteSpace($raw)) { continue }

      # MSI FileName may be in "short|long" format. Prefer the long name when present.
      $name = $raw
      if ($name.Contains("|")) {
        $parts = $name.Split("|")
        if ($parts.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($parts[1])) {
          $name = $parts[1]
        } elseif ($parts.Count -ge 1) {
          $name = $parts[0]
        }
      }
      $name = $name.Trim()
      if (-not [string]::IsNullOrWhiteSpace($name)) {
        $fileNames += $name
      }
    }

    $required = @("LICENSE", "NOTICE")
    $missing = @()
    foreach ($req in $required) {
      $found = $false
      foreach ($n in $fileNames) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($n)
        if ($base -ieq $req) { $found = $true; break }
      }
      if (-not $found) {
        $missing += $req
      }
    }

    if ($missing.Count -gt 0) {
      $presentSample = @($fileNames | Sort-Object -Unique | Select-Object -First 50) -join ", "
      throw "MSI installer is missing required compliance files: $($missing -join ", "). Expected LICENSE/NOTICE to be included in the installed app directory (typically under resources\\). File table sample: $presentSample"
    }
  }
 
  function Test-ByteArrayContainsSubsequence {
    param(
      [Parameter(Mandatory = $true)] [byte[]]$Haystack,
      [Parameter(Mandatory = $true)] [byte[]]$Needle
    )
 
    if ($Needle.Length -eq 0) { return $true }
    if ($Haystack.Length -lt $Needle.Length) { return $false }
 
    $first = $Needle[0]
    $limit = $Haystack.Length - $Needle.Length
    for ($i = 0; $i -le $limit; $i++) {
      if ($Haystack[$i] -ne $first) { continue }
      $match = $true
      for ($j = 1; $j -lt $Needle.Length; $j++) {
        if ($Haystack[$i + $j] -ne $Needle[$j]) { $match = $false; break }
      }
      if ($match) { return $true }
    }
    return $false
  }
 
  function Find-BinaryMarkerInFile {
    <#
      Streaming substring search over a binary file.
      Returns the marker string found, or $null.
 
      This mirrors the marker scan strategy used by scripts/ci/check-windows-webview2-installer.py.
    #>
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$File,
      [Parameter(Mandatory = $true)]
      [string[]]$MarkerStrings
    )
 
    $markers = @($MarkerStrings | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToString() } | Sort-Object -Unique)
    if ($markers.Count -eq 0) {
      return $null
    }
 
    $patterns = New-Object System.Collections.Generic.List[object]
    foreach ($m in $markers) {
      # Search for both UTF-8/ASCII and UTF-16LE encodings.
      $patterns.Add([pscustomobject]@{ Marker = $m; Bytes = [System.Text.Encoding]::UTF8.GetBytes($m) })
      $patterns.Add([pscustomobject]@{ Marker = $m; Bytes = [System.Text.Encoding]::Unicode.GetBytes($m) }) # UTF-16LE
    }
 
    $maxLen = 0
    foreach ($p in $patterns) {
      $len = $p.Bytes.Length
      if ($len -gt $maxLen) { $maxLen = $len }
    }
    $overlap = [Math]::Max(0, $maxLen - 1)
 
    $bufferSize = 1024 * 1024 # 1 MiB
    $buffer = New-Object byte[] $bufferSize
    $carry = New-Object byte[] 0
 
    $stream = [System.IO.File]::Open($File.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try {
      while ($true) {
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { break }
 
        $data = New-Object byte[] ($carry.Length + $read)
        if ($carry.Length -gt 0) {
          [Array]::Copy($carry, 0, $data, 0, $carry.Length)
        }
        [Array]::Copy($buffer, 0, $data, $carry.Length, $read)
 
        foreach ($p in $patterns) {
          if ($p.Bytes.Length -eq 0) { continue }
          if (Test-ByteArrayContainsSubsequence -Haystack $data -Needle $p.Bytes) {
            return $p.Marker
          }
        }
 
        if ($overlap -gt 0) {
          $carryLen = [Math]::Min($overlap, $data.Length)
          $carry = New-Object byte[] $carryLen
          if ($carryLen -gt 0) {
            [Array]::Copy($data, $data.Length - $carryLen, $carry, 0, $carryLen)
          }
        } else {
          $carry = New-Object byte[] 0
        }
      }
    } finally {
      $stream.Dispose()
    }
 
    return $null
  }
 
  function Find-ExeDesktopIntegrationMarker {
    <#
      Best-effort marker scan for desktop integration metadata in NSIS installers.
 
      We look for at least one of:
        - URL Protocol
        - x-scheme-handler/<scheme>
        - \<scheme>\shell\open\command
        - .xlsx
 
      NOTE: This validation is heuristic.
    #>
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Exe,
      [Parameter(Mandatory = $true)]
      [string]$ExtensionNoDot,
      [Parameter(Mandatory = $true)]
      [string]$UrlScheme
    )
 
    $dotExt = "." + ($ExtensionNoDot.Trim().TrimStart("."))
    $scheme = $UrlScheme.Trim()
    if ([string]::IsNullOrWhiteSpace($scheme)) { $scheme = "formula" }
    $scheme = $scheme.TrimEnd("/").TrimEnd(":")
 
      $markers = @(
        "URL Protocol",
        "URL protocol",
        # Token-delimited marker to avoid prefix false-positives (e.g. avoid treating
        # x-scheme-handler/<scheme>-extra as satisfying x-scheme-handler/<scheme>).
        "x-scheme-handler/$scheme;",
        "\$scheme\shell\open\command",
        "$scheme\shell\open\command",
        $dotExt,
        $dotExt.ToUpperInvariant()
      )
 
    return Find-BinaryMarkerInFile -File $Exe -MarkerStrings $markers
  }
 
  function Find-ExeUrlProtocolMarker {
    <#
      Best-effort marker scan for URL protocol registration in NSIS installers.
 
      NOTE: This validation is heuristic.
    #>
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Exe,
      [Parameter(Mandatory = $true)]
      [string]$UrlScheme
    )
 
    $scheme = $UrlScheme.Trim()
    if ([string]::IsNullOrWhiteSpace($scheme)) { $scheme = "formula" }
    $scheme = $scheme.TrimEnd("/").TrimEnd(":")
 
    $markers = @(
      # Token-delimited marker to avoid prefix false-positives (e.g. avoid treating
      # x-scheme-handler/<scheme>-extra as satisfying x-scheme-handler/<scheme>).
      "x-scheme-handler/$scheme;",
      "Software\\Classes\\$scheme\\shell\\open\\command",
      "Software\Classes\$scheme\shell\open\command",
      "HKEY_CLASSES_ROOT\\$scheme\\shell\\open\\command",
      "HKEY_CLASSES_ROOT\$scheme\shell\open\command",
      "HKCR\\$scheme\\shell\\open\\command",
      "HKCR\$scheme\shell\open\command",
      "\$scheme\shell\open\command",
      "$scheme\shell\open\command"
    )
 
    return Find-BinaryMarkerInFile -File $Exe -MarkerStrings $markers
  }

  function Test-ExeHasFileAssociationHints {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Exe,
      [Parameter(Mandatory = $true)]
      [string]$ExtensionNoDot
    )

    Write-Host "File association check (NSIS/EXE, best-effort): $($Exe.FullName)"

    # Best-effort validation: streaming scan for registry path strings that strongly suggest file
    # association registration. (We intentionally avoid relying on external NSIS parsing tools.)
    $dotExt = "." + ($ExtensionNoDot.Trim().TrimStart("."))
    $dotExtUpper = $dotExt.ToUpperInvariant()

    $strongNeedles = @(
      "Software\\Classes\\$dotExt",
      "Software\Classes\$dotExt",
      "HKEY_CLASSES_ROOT\\$dotExt",
      "HKEY_CLASSES_ROOT\$dotExt",
      "HKCR\\$dotExt",
      "HKCR $dotExt"
    )
    $strongMarker = Find-BinaryMarkerInFile -File $Exe -MarkerStrings $strongNeedles
    if (-not [string]::IsNullOrWhiteSpace($strongMarker)) {
      return $true
    }

    # Fallback: require both a registry-ish context string and the extension itself.
    $contextNeedles = @("Software\Classes", "HKEY_CLASSES_ROOT", "HKCR", "WriteRegStr", "OpenWithProgids")
    $contextMarker = Find-BinaryMarkerInFile -File $Exe -MarkerStrings $contextNeedles
    if ([string]::IsNullOrWhiteSpace($contextMarker)) {
      return $false
    }

    $extMarker = Find-BinaryMarkerInFile -File $Exe -MarkerStrings @($dotExt, $dotExtUpper)
    return (-not [string]::IsNullOrWhiteSpace($extMarker))
  }

  function Assert-ExeContainsComplianceArtifacts {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Exe,
      # When true, missing 7-Zip tooling will be treated as a warning instead of an error.
      [switch]$BestEffort
    )

    Write-Host "Compliance artifact check (NSIS/EXE): $($Exe.FullName)"

    $sevenZip = Get-SevenZipPath
    if ([string]::IsNullOrWhiteSpace($sevenZip)) {
      $msg = "7-Zip (7z) not found; cannot validate NSIS installer payload includes LICENSE/NOTICE. Install 7-Zip or ensure 7z.exe is on PATH."
      if ($BestEffort) {
        Write-Warning "$msg Skipping EXE compliance validation because MSI installers are present and are treated as authoritative."
        return
      }
      throw $msg
    }

    $tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("formula-nsis-extract-" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $tmpRoot | Out-Null

    try {
      # Extract the installer payload. 7z supports NSIS installers and is the most
      # reliable way to inspect the embedded file set without performing a full install.
      #
      # - `x` preserves directory structure when possible
      # - `-y` assumes Yes on all prompts (overwrite in temp dir)
      # - `-o<dir>` sets output directory
      & $sevenZip x "-o$tmpRoot" -y $Exe.FullName | Out-Null
      if ($LASTEXITCODE -ne 0) {
        throw "7z extraction failed for $($Exe.FullName) (exit code $LASTEXITCODE)."
      }

      $files = @(Get-ChildItem -LiteralPath $tmpRoot -Recurse -File -ErrorAction SilentlyContinue)
      $missing = @()
      foreach ($req in @("LICENSE", "NOTICE")) {
        # Allow extensions (e.g. LICENSE.txt) as long as the base name matches.
        $found = $files | Where-Object { $_.Name -ieq $req -or $_.BaseName -ieq $req } | Select-Object -First 1
        if (-not $found) {
          $missing += $req
        }
      }

      if ($missing.Count -gt 0) {
        throw "EXE installer payload is missing required compliance files: $($missing -join ", "). Expected LICENSE/NOTICE to be included in the installed app directory."
      }
    } finally {
      Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
  }

  $assocSpec = Get-ExpectedFileAssociationSpec -RepoRoot $repoRoot
  $expectedExtensions = @()
  if ($null -ne $assocSpec -and ($assocSpec.PSObject.Properties.Name -contains "Extensions")) {
    $expectedExtensions = @(
      $assocSpec.Extensions |
        Where-Object { $_ } |
        ForEach-Object { $_.ToString().Trim().TrimStart(".").ToLowerInvariant() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique
    )
  }
  if ($expectedExtensions.Count -eq 0) {
    $expectedExtensions = @("xlsx")
  }

  # MSI validation is authoritative; validate every configured file association extension.
  # For NSIS `.exe` marker scanning (best-effort), pick a stable representative extension (prefer `.xlsx`).
  $requiredExtensionNoDot = "xlsx"
  if (-not ($expectedExtensions -contains "xlsx")) {
    $requiredExtensionNoDot = ($expectedExtensions | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($requiredExtensionNoDot)) {
      $requiredExtensionNoDot = "xlsx"
    }
  }

  $mimeTypesByExtension = @{}
  if (
    $null -ne $assocSpec -and
    ($assocSpec.PSObject.Properties.Name -contains "MimeTypesByExtension") -and
    ($assocSpec.MimeTypesByExtension -is [hashtable])
  ) {
    $mimeTypesByExtension = $assocSpec.MimeTypesByExtension
  }

  $urlSpec = Get-ExpectedUrlProtocolSpec -RepoRoot $repoRoot
  $primaryScheme = ""
  $candidateSchemes = @()
  if ($null -ne $urlSpec -and ($urlSpec.PSObject.Properties.Name -contains "Schemes")) {
    $candidateSchemes = @($urlSpec.Schemes | Where-Object { $_ } | ForEach-Object { $_.ToString().Trim().TrimEnd("/").TrimEnd(":").ToLowerInvariant() } | Where-Object { $_ })
  }
  $expectedSchemes = @()
  if ($candidateSchemes.Count -gt 0) {
    $expectedSchemes = @($candidateSchemes | Sort-Object -Unique)
  } else {
    $expectedSchemes = @("formula")
  }
  # Prefer `formula` if it exists in the config (stable external deep link scheme) as the
  # primary one for best-effort validations (NSIS marker scan / MSI binary scan fallback).
  if ($expectedSchemes -contains "formula") {
    $primaryScheme = "formula"
  } else {
    $primaryScheme = ($expectedSchemes | Select-Object -First 1)
  }
  $primaryScheme = $primaryScheme.Trim().TrimEnd("/").TrimEnd(":")

  if ($msiInstallers.Count -gt 0) {
    foreach ($msi in $msiInstallers) {
      try {
        foreach ($ext in $expectedExtensions) {
          $expectedMime = ""
          if ($mimeTypesByExtension.ContainsKey($ext)) {
            $expectedMime = ($mimeTypesByExtension[$ext]).ToString().Trim()
          }
          Assert-MsiDeclaresFileAssociation -Msi $msi -ExtensionNoDot $ext -ExpectedMimeType $expectedMime
        }
      } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'Failed to open MSI for inspection') {
          Write-Warning "MSI inspection tooling is unavailable; falling back to best-effort string scan for file association metadata. Details: $msg"
          $dotExt = "." + $requiredExtensionNoDot
          $needles = @(
            "Software\\Classes\\$dotExt",
            "Software\Classes\$dotExt",
            "HKEY_CLASSES_ROOT\\$dotExt",
            "HKEY_CLASSES_ROOT\$dotExt",
            "HKCR\\$dotExt",
            "HKCR $dotExt"
          )
          if ($requiredExtensionNoDot -ieq "xlsx") {
            # Common ProgId used by Excel for `.xlsx` file associations.
            $needles += "Excel.Sheet.12"
          }
          $marker = Find-BinaryMarkerInFile -File $msi -MarkerStrings $needles
          if ([string]::IsNullOrWhiteSpace($marker)) {
            throw "Unable to inspect MSI tables AND did not find file-association-related strings for '.$requiredExtensionNoDot' in the MSI binary: $($msi.FullName)"
          }
          Write-Warning "MSI table inspection failed, but the MSI contained marker '$marker' related to '.$requiredExtensionNoDot'. Assuming file association metadata is present (best-effort)."
        } else {
          throw
        }
      }
 
      try {
        foreach ($scheme in $expectedSchemes) {
          Assert-MsiRegistersUrlProtocol -Msi $msi -Scheme $scheme
        }
      } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'Failed to open MSI for inspection') {
          Write-Warning "MSI inspection tooling is unavailable; falling back to best-effort string scan for URL protocol metadata. Details: $msg"
          foreach ($scheme in $expectedSchemes) {
            $schemeNeedles = @(
              # Prefer scheme-specific paths that include a separator after the scheme name to
              # avoid prefix false-positives (e.g. don't treat 'formula-extra' as satisfying 'formula').
              "\$scheme\shell\open\command",
              "$scheme\shell\open\command",
              "Software\\Classes\\$scheme\\shell\\open\\command",
              "Software\Classes\$scheme\shell\open\command",
              "HKEY_CLASSES_ROOT\\$scheme\\shell\\open\\command",
              "HKEY_CLASSES_ROOT\$scheme\shell\open\command",
              "HKCR\\$scheme\\shell\\open\\command",
              "HKCR\$scheme\shell\open\command"
            )
            $schemeMarker = Find-BinaryMarkerInFile -File $msi -MarkerStrings $schemeNeedles
            if ([string]::IsNullOrWhiteSpace($schemeMarker)) {
              throw "Unable to inspect MSI tables AND did not find URL-protocol-related strings for '$scheme://' in the MSI binary: $($msi.FullName)"
            }
            Write-Warning "MSI table inspection failed, but the MSI contained marker '$schemeMarker' related to '$scheme://'. Assuming URL protocol metadata is present (best-effort)."
          }
        } else {
          throw
        }
      }
      Assert-MsiContainsComplianceArtifacts -Msi $msi
    }
  } else {
    Write-Warning "No MSI installers found; falling back to best-effort EXE inspection for file association + URL protocol metadata."
  }

  if ($exeInstallers.Count -gt 0) {
    foreach ($exe in $exeInstallers) {
      $integrationMarker = Find-ExeDesktopIntegrationMarker -Exe $exe -ExtensionNoDot $requiredExtensionNoDot -UrlScheme $primaryScheme
      if ([string]::IsNullOrWhiteSpace($integrationMarker)) {
        throw "EXE installer did not contain any expected desktop integration marker strings. This validation is heuristic for NSIS installers.`n- Installer: $($exe.FullName)`n- Looked for: URL Protocol, x-scheme-handler/$primaryScheme;, \\$primaryScheme\\shell\\open\\command, .$requiredExtensionNoDot"
      }
      Write-Host ("EXE marker scan: found {0}" -f $integrationMarker)
 
      $ok = Test-ExeHasFileAssociationHints -Exe $exe -ExtensionNoDot $requiredExtensionNoDot
      if (-not $ok) {
        $msg = "EXE installer did not contain obvious file association registry strings for '.$requiredExtensionNoDot'."
        if ($msiInstallers.Count -gt 0) {
          Write-Warning "$msg MSI validation passed, so this is non-fatal. If users rely on the EXE installer, investigate NSIS file association wiring."
        } else {
          throw "$msg Without an MSI, we cannot reliably confirm Windows file associations are present."
        }
      }
 
      $protocolSchemes = @()
      if ($msiInstallers.Count -gt 0) {
        # MSI validation is authoritative; only scan the primary scheme for best-effort EXE validation.
        $protocolSchemes = @($primaryScheme)
      } else {
        $protocolSchemes = @($expectedSchemes)
      }
      foreach ($scheme in $protocolSchemes) {
        $protocolMarker = Find-ExeUrlProtocolMarker -Exe $exe -UrlScheme $scheme
        if ([string]::IsNullOrWhiteSpace($protocolMarker)) {
          $msg = "EXE installer did not contain obvious markers for '$scheme://' URL protocol registration (e.g. '$scheme\\shell\\open\\command'). This validation is heuristic for NSIS installers."
          if ($msiInstallers.Count -gt 0) {
            Write-Warning "$msg MSI validation passed, so this is non-fatal. If users rely on the EXE installer, investigate NSIS URL protocol wiring."
          } else {
            throw "$msg Without an MSI, we cannot reliably confirm Windows URL protocol registration is present."
          }
        } else {
          Write-Host ("EXE protocol marker scan ($scheme): found {0}" -f $protocolMarker)
        }
      }

      # NSIS is a distributed installer (alongside MSI). Ensure the same compliance artifacts
      # are present in its payload so users who install via EXE still receive LICENSE/NOTICE.
      Assert-ExeContainsComplianceArtifacts -Exe $exe -BestEffort:($msiInstallers.Count -gt 0)
    }
  }

  $expectedVersion = Get-ExpectedTauriVersion -RepoRoot $repoRoot
  Write-Host ("Expected desktop version (tauri.conf.json): {0}" -f $expectedVersion)
  $expectedProductName = Get-ExpectedProductName -RepoRoot $repoRoot
  if (-not [string]::IsNullOrWhiteSpace($expectedProductName)) {
    Write-Host ("Expected product name (tauri.conf.json productName): {0}" -f $expectedProductName)
  } else {
    Write-Warning "No productName found in tauri.conf.json. Skipping MSI ProductName validation."
  }
  $expectedUpgradeCode = Get-ExpectedWixUpgradeCode -RepoRoot $repoRoot
  if (-not [string]::IsNullOrWhiteSpace($expectedUpgradeCode)) {
    Write-Host ("Expected WiX UpgradeCode (tauri.conf.json): {0}" -f $expectedUpgradeCode)
  } else {
    Write-Warning "No bundle.windows.wix.upgradeCode found in tauri.conf.json. Skipping MSI UpgradeCode validation."
  }
  Write-Host ""

  foreach ($installer in $exeInstallers) {
    $vi = (Get-Item -LiteralPath $installer.FullName).VersionInfo
    $fileVersion = [string]$vi.FileVersion
    $productVersion = [string]$vi.ProductVersion

    # Accept if either FileVersion or ProductVersion matches (tooling varies by installer type).
    $ok = $false
    try {
      Assert-VersionMatch -ArtifactPath $installer.FullName -FoundVersion $fileVersion -ExpectedVersion $expectedVersion -Context "NSIS .exe FileVersion"
      $ok = $true
    } catch {
      # Ignore and try ProductVersion next.
    }
    if (-not $ok) {
      Assert-VersionMatch -ArtifactPath $installer.FullName -FoundVersion $productVersion -ExpectedVersion $expectedVersion -Context "NSIS .exe ProductVersion"
    }

    Write-Host ("version: OK (.exe) {0}" -f $installer.FullName)
  }

  foreach ($installer in $msiInstallers) {
    $msiVersion = Get-MsiProperty -MsiPath $installer.FullName -PropertyName "ProductVersion"
    if ($null -ne $msiVersion -and -not [string]::IsNullOrWhiteSpace([string]$msiVersion)) {
      Assert-VersionMatch -ArtifactPath $installer.FullName -FoundVersion ([string]$msiVersion) -ExpectedVersion $expectedVersion -Context "MSI ProductVersion"
      Write-Host ("version: OK (.msi) {0}" -f $installer.FullName)
    } else {
      Write-Warning ("Unable to read MSI ProductVersion for {0}. Skipping MSI version check because Windows Installer COM query failed. Consider enabling COM access or using an MSI inspection tool (lessmsi/msiinfo) in this environment." -f $installer.FullName)
    }

    if (-not [string]::IsNullOrWhiteSpace($expectedUpgradeCode)) {
      $msiUpgradeCode = Get-MsiProperty -MsiPath $installer.FullName -PropertyName "UpgradeCode"
      if ($null -ne $msiUpgradeCode -and -not [string]::IsNullOrWhiteSpace([string]$msiUpgradeCode)) {
        $expectedUpgradeNorm = Normalize-Guid -Value $expectedUpgradeCode
        $foundUpgradeNorm = Normalize-Guid -Value ([string]$msiUpgradeCode)
        if ([string]::IsNullOrWhiteSpace($expectedUpgradeNorm) -or [string]::IsNullOrWhiteSpace($foundUpgradeNorm)) {
          throw "Unable to parse MSI UpgradeCode GUID.`n- MSI: $($installer.FullName)`n- Expected (tauri.conf.json bundle.windows.wix.upgradeCode): $expectedUpgradeCode`n- Found (MSI UpgradeCode): $msiUpgradeCode"
        }
        if ($expectedUpgradeNorm -ne $foundUpgradeNorm) {
          throw "MSI UpgradeCode mismatch detected.`n- MSI: $($installer.FullName)`n- Expected (tauri.conf.json bundle.windows.wix.upgradeCode): $expectedUpgradeCode`n- Found (MSI UpgradeCode): $msiUpgradeCode"
        }
        Write-Host ("upgradeCode: OK (.msi) {0}" -f $installer.FullName)
      } else {
        Write-Warning ("Unable to read MSI UpgradeCode for {0}. Skipping MSI UpgradeCode check because Windows Installer COM query failed." -f $installer.FullName)
      }
    }

    if (-not [string]::IsNullOrWhiteSpace($expectedProductName)) {
      $msiProductName = Get-MsiProperty -MsiPath $installer.FullName -PropertyName "ProductName"
      if ($null -ne $msiProductName -and -not [string]::IsNullOrWhiteSpace([string]$msiProductName)) {
        $found = ([string]$msiProductName).Trim()
        if ($found -ne $expectedProductName) {
          throw "MSI ProductName mismatch detected.`n- MSI: $($installer.FullName)`n- Expected (tauri.conf.json productName): $expectedProductName`n- Found (MSI ProductName): $msiProductName"
        }
        Write-Host ("productName: OK (.msi) {0}" -f $installer.FullName)
      } else {
        Write-Warning ("Unable to read MSI ProductName for {0}. Skipping MSI ProductName check because Windows Installer COM query failed." -f $installer.FullName)
      }
    }
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
