[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,

    [Parameter(Mandatory = $true)]
    [string]$PreviewPath,

    [string]$OutputJsonPath = "",

    [string]$HeatmapPath = "",

    [int]$BlockSize = 32,

    [int]$TopRegions = 25
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path -LiteralPath $SourcePath)) { throw "Source image not found: $SourcePath" }
if (-not (Test-Path -LiteralPath $PreviewPath)) { throw "Preview image not found: $PreviewPath" }

function New-ResizedBitmap {
    param([System.Drawing.Image]$Image, [int]$Width, [int]$Height)
    $bmp = New-Object System.Drawing.Bitmap($Width, $Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($Image, 0, 0, $Width, $Height)
    $g.Dispose()
    return $bmp
}

function Get-Gray {
    param([System.Drawing.Color]$Color)
    return [int](0.299 * $Color.R + 0.587 * $Color.G + 0.114 * $Color.B)
}

$source = $null
$preview = $null
$sourceResized = $null
$heatmap = $null
$graphics = $null
try {
    $source = [System.Drawing.Image]::FromFile($SourcePath)
    $preview = New-Object System.Drawing.Bitmap($PreviewPath)
    $sourceResized = New-ResizedBitmap -Image $source -Width $preview.Width -Height $preview.Height

    $cols = [Math]::Ceiling($preview.Width / [double]$BlockSize)
    $rows = [Math]::Ceiling($preview.Height / [double]$BlockSize)
    $regions = New-Object System.Collections.ArrayList
    $totalDiff = 0.0
    $sampleCount = 0

    for ($row = 0; $row -lt $rows; $row++) {
        for ($col = 0; $col -lt $cols; $col++) {
            $x1 = $col * $BlockSize
            $y1 = $row * $BlockSize
            $x2 = [Math]::Min($x1 + $BlockSize, $preview.Width)
            $y2 = [Math]::Min($y1 + $BlockSize, $preview.Height)
            $sum = 0
            $count = 0
            for ($y = $y1; $y -lt $y2; $y += 2) {
                for ($x = $x1; $x -lt $x2; $x += 2) {
                    $s = $sourceResized.GetPixel($x, $y)
                    $p = $preview.GetPixel($x, $y)
                    $diff = [Math]::Abs((Get-Gray $s) - (Get-Gray $p))
                    $sum += $diff
                    $totalDiff += $diff
                    $count++
                    $sampleCount++
                }
            }
            if ($count -gt 0) {
                $avg = $sum / [double]$count
                if ($avg -gt 18) {
                    [void]$regions.Add([pscustomobject]@{
                        x = [int](($x1 + $x2) / 2)
                        y = [int](($y1 + $y2) / 2)
                        width = [int]($x2 - $x1)
                        height = [int]($y2 - $y1)
                        avg_gray_diff = [Math]::Round($avg, 2)
                    })
                }
            }
        }
    }

    $hotRegions = @($regions | Sort-Object avg_gray_diff -Descending | Select-Object -First $TopRegions)

    if (-not [string]::IsNullOrWhiteSpace($HeatmapPath)) {
        $heatmap = New-Object System.Drawing.Bitmap($preview.Width, $preview.Height)
        $graphics = [System.Drawing.Graphics]::FromImage($heatmap)
        $graphics.DrawImage($preview, 0, 0, $preview.Width, $preview.Height)
        foreach ($r in $hotRegions) {
            $alpha = [Math]::Min(180, [Math]::Max(45, [int]($r.avg_gray_diff * 2)))
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb($alpha, 255, 0, 0))
            $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(220, 180, 0, 0), 2)
            $rx = [int]($r.x - ($r.width / 2))
            $ry = [int]($r.y - ($r.height / 2))
            $graphics.FillRectangle($brush, $rx, $ry, [int]$r.width, [int]$r.height)
            $graphics.DrawRectangle($pen, $rx, $ry, [int]$r.width, [int]$r.height)
            $brush.Dispose()
            $pen.Dispose()
        }
        $heatmap.Save($HeatmapPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }

    $result = [pscustomobject]@{
        source = [System.IO.Path]::GetFullPath($SourcePath)
        preview = [System.IO.Path]::GetFullPath($PreviewPath)
        canvas = [pscustomobject]@{ width = $preview.Width; height = $preview.Height }
        mean_gray_diff = if ($sampleCount -gt 0) { [Math]::Round($totalDiff / [double]$sampleCount, 2) } else { 0 }
        block_size = $BlockSize
        hot_regions = $hotRegions
        heatmap = if ([string]::IsNullOrWhiteSpace($HeatmapPath)) { $null } else { [System.IO.Path]::GetFullPath($HeatmapPath) }
        note = "Diff is preview-space and visual only; inspect hot regions before accepting output."
    }

    $json = $result | ConvertTo-Json -Depth 8
    if ([string]::IsNullOrWhiteSpace($OutputJsonPath)) {
        $json
    }
    else {
        $json | Set-Content -LiteralPath $OutputJsonPath -Encoding UTF8
        $OutputJsonPath
    }
}
finally {
    if ($null -ne $graphics) { $graphics.Dispose() }
    if ($null -ne $heatmap) { $heatmap.Dispose() }
    if ($null -ne $sourceResized) { $sourceResized.Dispose() }
    if ($null -ne $preview) { $preview.Dispose() }
    if ($null -ne $source) { $source.Dispose() }
}
