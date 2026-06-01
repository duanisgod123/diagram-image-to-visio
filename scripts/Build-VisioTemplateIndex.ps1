[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TemplateDir,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$OutputJsonPath = "",

    [switch]$SkipExport
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path -LiteralPath $TemplateDir)) {
    throw "Template directory not found: $TemplateDir"
}
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

function Get-SafeName {
    param([string]$Name)
    return ($Name -replace '[^\p{L}\p{Nd}\.\-_]+', '_')
}

function Get-ImageSummary {
    param([string]$Path)
    $bmp = $null
    try {
        $bmp = New-Object System.Drawing.Bitmap($Path)
        $colors = @{}
        for ($y = 0; $y -lt $bmp.Height; $y += [Math]::Max([int]($bmp.Height / 80), 1)) {
            for ($x = 0; $x -lt $bmp.Width; $x += [Math]::Max([int]($bmp.Width / 80), 1)) {
                $c = $bmp.GetPixel($x, $y)
                if ($c.R -gt 245 -and $c.G -gt 245 -and $c.B -gt 245) { continue }
                $r = [int]([Math]::Floor($c.R / 32) * 32)
                $g = [int]([Math]::Floor($c.G / 32) * 32)
                $b = [int]([Math]::Floor($c.B / 32) * 32)
                $key = "#{0:X2}{1:X2}{2:X2}" -f $r, $g, $b
                if ($colors.ContainsKey($key)) { $colors[$key]++ } else { $colors[$key] = 1 }
            }
        }
        return [pscustomobject]@{
            width = $bmp.Width
            height = $bmp.Height
            aspect = [Math]::Round($bmp.Width / [double][Math]::Max($bmp.Height, 1), 3)
            colors = @($colors.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 8 | ForEach-Object {
                [pscustomobject]@{ color = $_.Key; samples = $_.Value }
            })
        }
    }
    finally {
        if ($null -ne $bmp) { $bmp.Dispose() }
    }
}

$files = @(Get-ChildItem -LiteralPath $TemplateDir -Filter "*.vsdx" -File)
$visio = $null
$items = New-Object System.Collections.ArrayList
try {
    if (-not $SkipExport) {
        $visio = New-Object -ComObject Visio.Application
        $visio.Visible = $false
    }

    foreach ($file in $files) {
        $safe = Get-SafeName -Name $file.BaseName
        $preview = Join-Path $OutputDir ($safe + ".png")
        $shapeCount = $null
        $pageCount = $null
        $errorMessage = $null
        try {
            if (-not $SkipExport) {
                $doc = $visio.Documents.Open($file.FullName)
                $pageCount = $doc.Pages.Count
                $page = $doc.Pages.Item(1)
                $shapeCount = $page.Shapes.Count
                $page.Export($preview)
                $doc.Close()
            }
            elseif (-not (Test-Path -LiteralPath $preview)) {
                $preview = ""
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            try { if ($null -ne $doc) { $doc.Close() } } catch {}
        }

        $summary = if ($preview -and (Test-Path -LiteralPath $preview)) { Get-ImageSummary -Path $preview } else { $null }
        [void]$items.Add([pscustomobject]@{
            name = $file.Name
            path = $file.FullName
            length = $file.Length
            preview = $preview
            pages = $pageCount
            shapes = $shapeCount
            image = $summary
            error = $errorMessage
        })
    }
}
finally {
    if ($null -ne $visio) { $visio.Quit() }
}

$index = [pscustomobject]@{
    template_dir = [System.IO.Path]::GetFullPath($TemplateDir)
    output_dir = [System.IO.Path]::GetFullPath($OutputDir)
    generated_at = (Get-Date).ToString("s")
    items = @($items)
}

$json = $index | ConvertTo-Json -Depth 8
if ([string]::IsNullOrWhiteSpace($OutputJsonPath)) {
    $json
}
else {
    $json | Set-Content -LiteralPath $OutputJsonPath -Encoding UTF8
    $OutputJsonPath
}
