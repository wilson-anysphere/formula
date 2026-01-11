<#
.SYNOPSIS
  Export chart images from Excel fixture workbooks as golden PNGs.

.DESCRIPTION
  Opens each workbook under `fixtures/charts/xlsx/` in Microsoft Excel (COM automation),
  finds the first embedded chart, sizes it, and exports it as a PNG under
  `fixtures/charts/golden/excel/`.

  This is Windows-only and requires Microsoft Excel desktop installed.

.PARAMETER FixturesDir
  Directory containing `.xlsx` chart fixtures.

.PARAMETER OutDir
  Output directory for golden PNGs.

.PARAMETER WidthPx
  Desired output width in pixels.

.PARAMETER HeightPx
  Desired output height in pixels.

.PARAMETER Visible
  Make Excel visible while running (useful for debugging).
#>

[CmdletBinding()]
param(
  [string]$FixturesDir = "fixtures/charts/xlsx",
  [string]$OutDir = "fixtures/charts/golden/excel",
  [int]$WidthPx = 800,
  [int]$HeightPx = 600,
  [switch]$Visible
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

function Get-PixelsPerInch {
  param(
    [Parameter(Mandatory = $true)]
    [object]$Excel,
    [switch]$X
  )

  try {
    if ($null -ne $Excel.ActiveWindow) {
      if ($X) {
        return [double]$Excel.ActiveWindow.PointsToScreenPixelsX(72)
      } else {
        return [double]$Excel.ActiveWindow.PointsToScreenPixelsY(72)
      }
    }
  } catch {}

  # Fallback: assume 96 DPI (common Windows setting).
  return 96.0
}

function PixelsToPoints {
  param(
    [Parameter(Mandatory = $true)]
    [int]$Pixels,
    [Parameter(Mandatory = $true)]
    [double]$PixelsPerInch
  )
  # 1 inch = 72 points.
  return [double]$Pixels * 72.0 / $PixelsPerInch
}

if (-not (Test-Path -LiteralPath $FixturesDir)) {
  throw "FixturesDir not found: $FixturesDir"
}

$fullOutDir = [System.IO.Path]::GetFullPath($OutDir)
if (-not (Test-Path -LiteralPath $fullOutDir)) {
  New-Item -ItemType Directory -Force -Path $fullOutDir | Out-Null
}

$excel = $null
$workbook = $null

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

  $files = Get-ChildItem -LiteralPath $FixturesDir -Filter *.xlsx | Sort-Object Name
  foreach ($file in $files) {
    $inPath = $file.FullName
    $stem = $file.BaseName
    $outPath = Join-Path $fullOutDir "$stem.png"

    Write-Host "Exporting $($file.Name) -> $([System.IO.Path]::GetFileName($outPath))"

    $workbook = $excel.Workbooks.Open($inPath, $false, $true) # read-only
    $chartObject = $null
    $chartObjects = $null
    $sheet = $null

    try {
      # Resolve DPI conversion after opening the workbook (ActiveWindow exists).
      $pxPerInchX = Get-PixelsPerInch -Excel $excel -X
      $pxPerInchY = Get-PixelsPerInch -Excel $excel
      $widthPt = PixelsToPoints -Pixels $WidthPx -PixelsPerInch $pxPerInchX
      $heightPt = PixelsToPoints -Pixels $HeightPx -PixelsPerInch $pxPerInchY

      foreach ($ws in @($workbook.Worksheets)) {
        $sheet = $ws
        try {
          $chartObjects = $sheet.ChartObjects()
          if ($chartObjects.Count -gt 0) {
            $chartObject = $chartObjects.Item(1)
            break
          }
        } catch {
          # Sheets without embedded charts will throw here; ignore.
        } finally {
          Release-ComObject $chartObjects
          $chartObjects = $null
          Release-ComObject $sheet
          $sheet = $null
        }
      }

      if ($null -eq $chartObject) {
        Write-Warning "No embedded chart found in $($file.Name); skipping"
        continue
      }

      # Resize in points before exporting.
      $chartObject.Width = $widthPt
      $chartObject.Height = $heightPt

      # Export to PNG.
      $chartObject.Chart.Export($outPath, "PNG") | Out-Null
    } finally {
      Release-ComObject $chartObject
      if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
        $workbook = $null
      }
    }
  }
} finally {
  if ($null -ne $excel) {
    try { $excel.Quit() } catch {}
  }
  Release-ComObject $excel
  $excel = $null

  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

