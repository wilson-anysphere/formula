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

function Read-UInt32BE {
  param(
    [Parameter(Mandatory = $true)]
    [byte[]]$Bytes,
    [Parameter(Mandatory = $true)]
    [int]$Offset
  )

  if ($Offset -lt 0 -or ($Offset + 4) -gt $Bytes.Length) {
    throw "Read-UInt32BE: out of bounds (offset=$Offset, len=$($Bytes.Length))"
  }

  return (
    ([uint32]$Bytes[$Offset]   -shl 24) -bor
    ([uint32]$Bytes[$Offset+1] -shl 16) -bor
    ([uint32]$Bytes[$Offset+2] -shl 8)  -bor
    ([uint32]$Bytes[$Offset+3])
  )
}

function Get-PngDimensions {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  $bytes = [System.IO.File]::ReadAllBytes($Path)
  if ($bytes.Length -lt 24) {
    throw "invalid PNG (too small): $Path"
  }

  # PNG signature.
  $sig = [byte[]](0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
  for ($i = 0; $i -lt $sig.Length; $i++) {
    if ($bytes[$i] -ne $sig[$i]) {
      throw "invalid PNG signature: $Path"
    }
  }

  $ihdrLen = Read-UInt32BE -Bytes $bytes -Offset 8
  if ($ihdrLen -lt 8) {
    throw "invalid PNG IHDR length ($ihdrLen): $Path"
  }
  $chunkType = [System.Text.Encoding]::ASCII.GetString($bytes[12..15])
  if ($chunkType -ne "IHDR") {
    throw "invalid PNG (first chunk is not IHDR): $Path"
  }

  $w = Read-UInt32BE -Bytes $bytes -Offset 16
  $h = Read-UInt32BE -Bytes $bytes -Offset 20
  return @($w, $h)
}

function Get-PngSampleUniqueColorCount {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [int]$Step = 40,
    [int]$MaxUnique = 20
  )

  try {
    Add-Type -AssemblyName System.Drawing | Out-Null
    $bmp = [System.Drawing.Bitmap]::new($Path)
    try {
      $set = New-Object System.Collections.Generic.HashSet[string]
      for ($y = 0; $y -lt $bmp.Height; $y += $Step) {
        for ($x = 0; $x -lt $bmp.Width; $x += $Step) {
          $c = $bmp.GetPixel($x, $y)
          [void]$set.Add("$($c.R),$($c.G),$($c.B)")
          if ($set.Count -ge $MaxUnique) {
            return $set.Count
          }
        }
      }
      return $set.Count
    } finally {
      $bmp.Dispose()
    }
  } catch {
    # Best-effort: if bitmap loading isn't available (e.g. missing GDI+), skip.
    return $null
  }
}

function Validate-GoldenPng {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [int]$WidthPx,
    [Parameter(Mandatory = $true)]
    [int]$HeightPx
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "Excel export did not create output PNG: $Path"
  }

  $dims = Get-PngDimensions -Path $Path
  $w = $dims[0]
  $h = $dims[1]
  if ($w -ne $WidthPx -or $h -ne $HeightPx) {
    throw "exported PNG has wrong size ($w x $h); expected ${WidthPx}x${HeightPx}: $Path"
  }

  $uniq = Get-PngSampleUniqueColorCount -Path $Path
  if ($null -ne $uniq -and $uniq -le 2) {
    Write-Warning "Exported PNG appears to be a placeholder/blank image (sample unique colors=$uniq): $Path"
  }
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
      Validate-GoldenPng -Path $outPath -WidthPx $WidthPx -HeightPx $HeightPx
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
