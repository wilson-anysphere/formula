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

  # Validate file association metadata is present in the produced installers.
  #
  # On Windows, `.xlsx` file associations are typically registered via MSI tables
  # (Extension/ProgId/Verb). This is the most reliable thing to validate in CI.
  #
  # For NSIS `.exe` installers, reliable inspection tooling is not always available on
  # GitHub-hosted runners. We do a best-effort string scan for registry paths that
  # indicate file association registration. If an MSI is present, MSI validation is
  # authoritative; the EXE scan is treated as a warning.
  function Get-ExpectedFileAssociationSpec {
    param(
      [Parameter(Mandatory = $true)]
      [string]$RepoRoot
    )

    $configPath = Join-Path $RepoRoot "apps/desktop/src-tauri/tauri.conf.json"
    $default = [pscustomobject]@{
      Extensions = @("xlsx")
      XlsxMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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

    $xlsxEntry = @(
      $fileAssociations |
        Where-Object {
          ($_.PSObject.Properties.Name -contains "ext") -and ($_.ext -contains "xlsx" -or $_.ext -contains ".xlsx")
        }
    ) | Select-Object -First 1
    $mime = $default.XlsxMimeType
    if (
      $null -ne $xlsxEntry -and
      ($xlsxEntry.PSObject.Properties.Name -contains "mimeType") -and
      -not [string]::IsNullOrWhiteSpace($xlsxEntry.mimeType)
    ) {
      $mime = ($xlsxEntry.mimeType).ToString()
    }

    return [pscustomobject]@{
      Extensions = @("xlsx")
      XlsxMimeType = $mime
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

  function Assert-MsiDeclaresFileAssociation {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Msi,
      [Parameter(Mandatory = $true)]
      [string]$ExtensionNoDot
    )

    Write-Host "File association check (MSI): $($Msi.FullName)"

    $tables = @()
    try {
      $tables = Get-MsiTableNames -MsiPath $Msi.FullName
    } catch {
      throw "Failed to open MSI for inspection: $($Msi.FullName)`n$($_.Exception.Message)"
    }

    if (-not ($tables -contains "Extension")) {
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
      throw "MSI Extension table row for '$ExtensionNoDot' exists but ProgId_ is empty. This suggests file association wiring is incomplete."
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

  function Test-StringContainsIgnoreCase {
    param(
      [Parameter(Mandatory = $true)] [string]$Haystack,
      [Parameter(Mandatory = $true)] [string]$Needle
    )
    return $Haystack.IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
  }

  function Test-ExeHasFileAssociationHints {
    param(
      [Parameter(Mandatory = $true)]
      [System.IO.FileInfo]$Exe,
      [Parameter(Mandatory = $true)]
      [string]$ExtensionNoDot
    )

    Write-Host "File association check (NSIS/EXE, best-effort): $($Exe.FullName)"

    # TODO: Replace this with a structured NSIS inspection (if/when reliable tooling becomes
    # available on GH runners). For now, scan for registry path strings that strongly suggest
    # file association registration.
    $dotExt = "." + $ExtensionNoDot
    $strongNeedles = @(
      "Software\\Classes\\$dotExt",
      "Software\Classes\$dotExt",
      "HKEY_CLASSES_ROOT\\$dotExt",
      "HKEY_CLASSES_ROOT\$dotExt",
      "HKCR\\$dotExt",
      "HKCR $dotExt"
    )
    $contextNeedles = @("Software\Classes", "HKEY_CLASSES_ROOT", "HKCR", "WriteRegStr", "OpenWithProgids")

    $bytes = [System.IO.File]::ReadAllBytes($Exe.FullName)
    $ascii = [System.Text.Encoding]::ASCII.GetString($bytes)
    $unicode = [System.Text.Encoding]::Unicode.GetString($bytes)

    foreach ($n in $strongNeedles) {
      if (Test-StringContainsIgnoreCase -Haystack $ascii -Needle $n) { return $true }
      if (Test-StringContainsIgnoreCase -Haystack $unicode -Needle $n) { return $true }
    }

    $hasContext =
      ($contextNeedles | Where-Object { Test-StringContainsIgnoreCase -Haystack $ascii -Needle $_ }).Count -gt 0 -or
      ($contextNeedles | Where-Object { Test-StringContainsIgnoreCase -Haystack $unicode -Needle $_ }).Count -gt 0
    $hasExt =
      (Test-StringContainsIgnoreCase -Haystack $ascii -Needle $dotExt) -or
      (Test-StringContainsIgnoreCase -Haystack $unicode -Needle $dotExt)

    return ($hasContext -and $hasExt)
  }

  $assocSpec = Get-ExpectedFileAssociationSpec -RepoRoot $repoRoot
  $requiredExtension = ($assocSpec.Extensions | Select-Object -First 1)
  if ([string]::IsNullOrWhiteSpace($requiredExtension)) {
    $requiredExtension = "xlsx"
  }
  $requiredExtensionNoDot = $requiredExtension.Trim().TrimStart(".")

  if ($msiInstallers.Count -gt 0) {
    foreach ($msi in $msiInstallers) {
      Assert-MsiDeclaresFileAssociation -Msi $msi -ExtensionNoDot $requiredExtensionNoDot
    }
  } else {
    Write-Warning "No MSI installers found; falling back to best-effort EXE inspection for file association metadata."
  }

  if ($exeInstallers.Count -gt 0) {
    foreach ($exe in $exeInstallers) {
      $ok = Test-ExeHasFileAssociationHints -Exe $exe -ExtensionNoDot $requiredExtensionNoDot
      if (-not $ok) {
        $msg = "EXE installer did not contain obvious file association registry strings for '.$requiredExtensionNoDot'."
        if ($msiInstallers.Count -gt 0) {
          Write-Warning "$msg MSI validation passed, so this is non-fatal. If users rely on the EXE installer, investigate NSIS file association wiring."
        } else {
          throw "$msg Without an MSI, we cannot reliably confirm Windows file associations are present."
        }
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
