[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ImagePath,

    [string]$OutputPath = "",

    [int]$Stride = 4,

    [int]$MinComponentPixels = 120
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path -LiteralPath $ImagePath)) {
    throw "Image file not found: $ImagePath"
}

function Test-ContentPixel {
    param([System.Drawing.Color]$Color)
    if ($Color.A -lt 20) { return $false }
    $max = [Math]::Max($Color.R, [Math]::Max($Color.G, $Color.B))
    $min = [Math]::Min($Color.R, [Math]::Min($Color.G, $Color.B))
    $isWhite = ($Color.R -gt 242 -and $Color.G -gt 242 -and $Color.B -gt 242)
    $isLightGray = ($Color.R -gt 225 -and $Color.G -gt 225 -and $Color.B -gt 225 -and ($max - $min) -lt 18)
    return (-not ($isWhite -or $isLightGray))
}

function Add-ColorCount {
    param($Map, [System.Drawing.Color]$Color)
    if (-not (Test-ContentPixel -Color $Color)) { return }
    $r = [int]([Math]::Floor($Color.R / 32) * 32)
    $g = [int]([Math]::Floor($Color.G / 32) * 32)
    $b = [int]([Math]::Floor($Color.B / 32) * 32)
    $key = "#{0:X2}{1:X2}{2:X2}" -f $r, $g, $b
    if ($Map.ContainsKey($key)) { $Map[$key]++ } else { $Map[$key] = 1 }
}

$bitmap = $null
try {
    $bitmap = New-Object System.Drawing.Bitmap($ImagePath)
    $width = $bitmap.Width
    $height = $bitmap.Height
    $gridW = [Math]::Ceiling($width / [double]$Stride)
    $gridH = [Math]::Ceiling($height / [double]$Stride)
    $mask = New-Object 'bool[]' ($gridW * $gridH)
    $colorCounts = @{}
    $minX = $width; $minY = $height; $maxX = -1; $maxY = -1

    for ($gy = 0; $gy -lt $gridH; $gy++) {
        for ($gx = 0; $gx -lt $gridW; $gx++) {
            $x = [Math]::Min($gx * $Stride, $width - 1)
            $y = [Math]::Min($gy * $Stride, $height - 1)
            $c = $bitmap.GetPixel($x, $y)
            Add-ColorCount -Map $colorCounts -Color $c
            if (Test-ContentPixel -Color $c) {
                $mask[$gy * $gridW + $gx] = $true
                if ($x -lt $minX) { $minX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
    }

    $visited = New-Object 'bool[]' ($gridW * $gridH)
    $components = New-Object System.Collections.ArrayList
    $queue = New-Object System.Collections.Generic.Queue[int]
    for ($i = 0; $i -lt $mask.Count; $i++) {
        if (-not $mask[$i] -or $visited[$i]) { continue }
        $visited[$i] = $true
        $queue.Enqueue($i)
        $count = 0
        $cMinGX = $gridW; $cMinGY = $gridH; $cMaxGX = 0; $cMaxGY = 0
        while ($queue.Count -gt 0) {
            $idx = $queue.Dequeue()
            $gx = $idx % $gridW
            $gy = [Math]::Floor($idx / $gridW)
            $count++
            if ($gx -lt $cMinGX) { $cMinGX = $gx }
            if ($gy -lt $cMinGY) { $cMinGY = $gy }
            if ($gx -gt $cMaxGX) { $cMaxGX = $gx }
            if ($gy -gt $cMaxGY) { $cMaxGY = $gy }
            foreach ($delta in @(@(-1,0), @(1,0), @(0,-1), @(0,1))) {
                $nx = $gx + $delta[0]
                $ny = $gy + $delta[1]
                if ($nx -lt 0 -or $ny -lt 0 -or $nx -ge $gridW -or $ny -ge $gridH) { continue }
                $nidx = $ny * $gridW + $nx
                if ($mask[$nidx] -and -not $visited[$nidx]) {
                    $visited[$nidx] = $true
                    $queue.Enqueue($nidx)
                }
            }
        }
        if ($count -ge $MinComponentPixels) {
            $x1 = $cMinGX * $Stride
            $y1 = $cMinGY * $Stride
            $x2 = [Math]::Min((($cMaxGX + 1) * $Stride), $width)
            $y2 = [Math]::Min((($cMaxGY + 1) * $Stride), $height)
            [void]$components.Add([pscustomobject]@{
                x = [int](($x1 + $x2) / 2)
                y = [int](($y1 + $y2) / 2)
                width = [int]($x2 - $x1)
                height = [int]($y2 - $y1)
                pixel_count = [int]($count * $Stride * $Stride)
            })
        }
    }

    $dominantColors = $colorCounts.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 12 |
        ForEach-Object { [pscustomobject]@{ color = $_.Key; samples = $_.Value } }

    $result = [pscustomobject]@{
        image = [System.IO.Path]::GetFullPath($ImagePath)
        canvas = [pscustomobject]@{ width = $width; height = $height }
        content_bbox = if ($maxX -ge 0) {
            [pscustomobject]@{
                x = [int](($minX + $maxX) / 2)
                y = [int](($minY + $maxY) / 2)
                width = [int]($maxX - $minX + 1)
                height = [int]($maxY - $minY + 1)
            }
        } else { $null }
        dominant_colors = @($dominantColors)
        components = @($components | Sort-Object pixel_count -Descending | Select-Object -First 80)
        analysis = [pscustomobject]@{
            stride = $Stride
            min_component_pixels = $MinComponentPixels
            note = "Connected components are visual guidance only; use OCR/layout tools when available."
        }
    }

    $json = $result | ConvertTo-Json -Depth 8
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $json
    }
    else {
        $json | Set-Content -LiteralPath $OutputPath -Encoding UTF8
        $OutputPath
    }
}
finally {
    if ($null -ne $bitmap) { $bitmap.Dispose() }
}
