[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [ValidateSet("Auto", "Landscape", "Portrait")]
    [string]$Orientation = "Auto",

    [ValidateSet("Auto", "A4", "A3", "Letter", "Legal")]
    [string]$PageSize = "Auto",

    [bool]$PreserveColor = $true,

    [string]$UncertainConnectorColor = "#FF00FF",

    [bool]$Visible = $false
)

$ErrorActionPreference = "Stop"

function Test-OutputFileWritable {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    }
    catch {
        throw "Output file is locked or not writable: $Path. Close it in Visio or choose a different output path."
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function ConvertTo-RgbFormula {
    param(
        [AllowNull()]
        [string]$Color,
        [string]$Fallback = "#000000"
    )

    $value = if ([string]::IsNullOrWhiteSpace($Color)) { $Fallback } else { $Color.Trim() }
    if ($value -notmatch "^#?(?<hex>[0-9A-Fa-f]{6})$") {
        $value = $Fallback
    }

    $hex = ($value -replace "^#", "")
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return "RGB($r,$g,$b)"
}

function Set-FormulaIfPresent {
    param(
        $Shape,
        [string]$CellName,
        [string]$Formula
    )

    if ([string]::IsNullOrWhiteSpace($Formula)) {
        return
    }

    try {
        $Shape.CellsU($CellName).FormulaU = $Formula
    }
    catch {
    }
}

function Set-ResultIfPresent {
    param(
        $Shape,
        [string]$CellName,
        [double]$Value
    )

    try {
        $Shape.CellsU($CellName).ResultIU = $Value
    }
    catch {
    }
}

function Get-LinePatternCode {
    param([string]$Pattern)

    $normalized = if ($null -eq $Pattern) { "" } else { $Pattern.ToLowerInvariant() }
    switch ($normalized) {
        "dashed" { return 2 }
        "dotted" { return 3 }
        default { return 1 }
    }
}

function Get-EndArrowCode {
    param([string]$Arrow)

    $normalized = if ($null -eq $Arrow) { "" } else { $Arrow.ToLowerInvariant() }
    switch ($normalized) {
        "open" { return 4 }
        "none" { return 0 }
        default { return 13 }
    }
}

function Get-ArrowSizeCode {
    param([string]$ArrowSize)

    $normalized = if ($null -eq $ArrowSize) { "" } else { $ArrowSize.ToLowerInvariant() }
    switch ($normalized) {
        "very_small" { return 0 }
        "tiny" { return 0 }
        "small" { return 1 }
        "medium" { return 2 }
        "large" { return 3 }
        "very_large" { return 4 }
        "huge" { return 5 }
        default { return 1 }
    }
}

function Resolve-FontId {
    param([string]$FontFamily)

    if ([string]::IsNullOrWhiteSpace($FontFamily)) {
        return $null
    }

    $doc = $script:VisioDocument
    if ($null -eq $doc) {
        return $null
    }

    try {
        return [int]$doc.Fonts.ItemU($FontFamily).ID
    }
    catch {
        return $null
    }
}

function Resolve-Orientation {
    param($Spec)

    if ($Orientation -ne "Auto") {
        return $Orientation
    }

    $specOrientation = $Spec.page.orientation
    if ($specOrientation -and @("Landscape", "Portrait") -contains $specOrientation) {
        return $specOrientation
    }

    if ([double]$Spec.canvas.width -ge [double]$Spec.canvas.height) {
        return "Landscape"
    }

    return "Portrait"
}

function Resolve-PageDimensions {
    param([string]$ResolvedOrientation)

    $sizeName = if ($PageSize -eq "Auto") {
        if ($Spec.page.page_size) { [string]$Spec.page.page_size } else { "A4" }
    }
    else {
        $PageSize
    }

    switch ($sizeName) {
        "A3" { $width = 16.54; $height = 11.69 }
        "Letter" { $width = 11.0; $height = 8.5 }
        "Legal" { $width = 14.0; $height = 8.5 }
        default { $width = 11.69; $height = 8.27 }
    }

    if ($ResolvedOrientation -eq "Portrait" -and $width -gt $height) {
        return @{ Width = $height; Height = $width }
    }

    if ($ResolvedOrientation -eq "Landscape" -and $height -gt $width) {
        return @{ Width = $height; Height = $width }
    }

    return @{ Width = $width; Height = $height }
}

function Get-PageMargin {
    param($Spec)

    $margin = 0.35
    if ($Spec.page -and $null -ne $Spec.page.margin) {
        try {
            $margin = [double]$Spec.page.margin
        }
        catch {
            $margin = 0.35
        }
    }

    return [Math]::Max([Math]::Min($margin, 1.5), 0.0)
}

function New-CoordinateMapper {
    param(
        [double]$CanvasWidth,
        [double]$CanvasHeight,
        [double]$PageWidth,
        [double]$PageHeight,
        [double]$Margin = 0.35
    )

    $margin = $Margin
    $usableWidth = [Math]::Max($PageWidth - (2 * $margin), 1.0)
    $usableHeight = [Math]::Max($PageHeight - (2 * $margin), 1.0)
    $scale = [Math]::Min($usableWidth / [Math]::Max($CanvasWidth, 1.0), $usableHeight / [Math]::Max($CanvasHeight, 1.0))
    $offsetX = $margin + (($usableWidth - ($CanvasWidth * $scale)) / 2.0)
    $offsetY = $margin + (($usableHeight - ($CanvasHeight * $scale)) / 2.0)

    return @{
        Scale = $scale
        Margin = $margin
        OffsetX = $offsetX
        OffsetY = $offsetY
        CanvasHeight = $CanvasHeight
    }
}

function Convert-ToPageX {
    param($Map, [double]$X)
    return $Map.OffsetX + ($X * $Map.Scale)
}

function Convert-ToPageY {
    param($Map, [double]$Y)
    return $Map.OffsetY + (($Map.CanvasHeight - $Y) * $Map.Scale)
}

function Convert-ToPageWidth {
    param($Map, [double]$Width)
    return [Math]::Max($Width * $Map.Scale, 0.08)
}

function Convert-ToPageHeight {
    param($Map, [double]$Height)
    return [Math]::Max($Height * $Map.Scale, 0.08)
}

function Convert-PointToPage {
    param($Map, [double]$X, [double]$Y)
    return @(
        (Convert-ToPageX -Map $Map -X $X),
        (Convert-ToPageY -Map $Map -Y $Y)
    )
}

function Get-TextPaddingValue {
    param(
        $Padding,
        [string]$Side
    )

    if ($null -eq $Padding) {
        return 0.0
    }

    $prop = $Padding.PSObject.Properties[$Side]
    if ($null -eq $prop) {
        return 0.0
    }

    return [double]$prop.Value
}

function Resolve-FontFamily {
    param(
        $Item,
        $Defaults
    )

    if ($Item.font_family) {
        return [string]$Item.font_family
    }

    if ($Defaults.font_family) {
        return [string]$Defaults.font_family
    }

    $text = if ($Item.text) { [string]$Item.text } else { "" }
    if ($text -match '[\u3400-\u9FFF]') {
        return "Microsoft YaHei"
    }

    return "Arial"
}

function Get-TextMetrics {
    param([string]$Text)

    $raw = if ($null -eq $Text) { "" } else { $Text }
    $normalized = $raw -replace "`r", ""
    $lines = @($normalized -split "`n")
    if ($lines.Count -eq 0) {
        $lines = @("")
    }

    $maxUnits = 0.0
    foreach ($line in $lines) {
        $units = 0.0
        foreach ($ch in $line.ToCharArray()) {
            $charCode = [int][char]$ch
            if ([char]::IsWhiteSpace($ch)) {
                $units += 0.35
            }
            elseif ($charCode -ge 0x3400 -and $charCode -le 0x9FFF) {
                $units += 1.0
            }
            elseif ([char]::IsUpper($ch)) {
                $units += 0.72
            }
            elseif ([char]::IsLower($ch) -or [char]::IsDigit($ch)) {
                $units += 0.58
            }
            else {
                $units += 0.52
            }
        }

        if ($units -gt $maxUnits) {
            $maxUnits = $units
        }
    }

    return @{
        LineCount = [Math]::Max($lines.Count, 1)
        MaxUnits = [Math]::Max($maxUnits, 1.0)
    }
}

function Get-LuminanceScore {
    param([string]$Color)

    if ([string]::IsNullOrWhiteSpace($Color) -or $Color -notmatch "^#?(?<hex>[0-9A-Fa-f]{6})$") {
        return 255.0
    }

    $hex = ($Color -replace "^#", "")
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return ((0.299 * $r) + (0.587 * $g) + (0.114 * $b))
}

function Resolve-RequestedFontSize {
    param(
        $Item,
        $Defaults
    )

    if ($Item.font_size) {
        return [double]$Item.font_size
    }

    $text = if ($Item.text) { [string]$Item.text } else { "" }
    $width = if ($null -ne $Item.width) { [double]$Item.width } else { 0.0 }
    $height = if ($null -ne $Item.height) { [double]$Item.height } else { 0.0 }
    $fillColor = if ($Item.fill_color) { [string]$Item.fill_color } elseif ($Defaults.fill_color) { [string]$Defaults.fill_color } else { "#FFFFFF" }
    $textColor = if ($Item.text_color) { [string]$Item.text_color } elseif ($Defaults.text_color) { [string]$Defaults.text_color } else { "#111111" }
    $fillLum = Get-LuminanceScore -Color $fillColor
    $textLum = Get-LuminanceScore -Color $textColor
    $inverseContrast = ($fillLum -lt 150.0 -and $textLum -gt 180.0)
    $isShort = (($text -replace '\s','').Length -le 12)

    if ($width -ge 700 -and $height -ge 36) {
        return 16.0
    }

    if (($Item.id -like '*title*' -or $Item.id -like 'header' -or $Item.id -like 'footer') -and $height -ge 20) {
        return 15.0
    }

    if ($inverseContrast -and $height -ge 28 -and $isShort) {
        return 11.0
    }

    if ($width -le 90 -and $height -ge 50) {
        return 7.0
    }

    if ($width -le 90 -and $height -lt 50) {
        return 9.0
    }

    if ($height -ge 30) {
        return 10.0
    }

    if ($Defaults.font_size) {
        return [double]$Defaults.font_size
    }

    return 10.0
}

function Resolve-FontSize {
    param(
        $Item,
        $Defaults,
        $Map,
        [double]$RequestedFontSize,
        [bool]$IsConnector
    )

    if ($IsConnector -or -not $Map -or -not $Item.text -or -not $Item.width -or -not $Item.height) {
        return $RequestedFontSize
    }

    $fontSizeMode = if ($Item.font_size_mode) { [string]$Item.font_size_mode } elseif ($Defaults.font_size_mode) { [string]$Defaults.font_size_mode } else { "fit" }
    if ($fontSizeMode.ToLowerInvariant() -eq "fixed") {
        return $RequestedFontSize
    }

    $textPadding = if ($Item.text_padding) { $Item.text_padding } elseif ($Defaults.text_padding) { $Defaults.text_padding } else { $null }
    $padLeft = Get-TextPaddingValue -Padding $textPadding -Side "left"
    $padRight = Get-TextPaddingValue -Padding $textPadding -Side "right"
    $padTop = Get-TextPaddingValue -Padding $textPadding -Side "top"
    $padBottom = Get-TextPaddingValue -Padding $textPadding -Side "bottom"

    $availableWidth = [Math]::Max(([double]$Item.width - $padLeft - $padRight), 10.0)
    $availableHeight = [Math]::Max(([double]$Item.height - $padTop - $padBottom), 8.0)
    $metrics = Get-TextMetrics -Text ([string]$Item.text)

    $heightFactor = if ($metrics.LineCount -eq 1) { 1.30 } else { 1.55 }
    $maxByWidth = $availableWidth / [Math]::Max($metrics.MaxUnits, 1.0)
    $maxByHeight = $availableHeight / ([Math]::Max($metrics.LineCount, 1) * $heightFactor)
    $resolved = [Math]::Min($RequestedFontSize, [Math]::Min($maxByWidth, $maxByHeight))

    return [Math]::Max([Math]::Round($resolved, 2), 7.0)
}

function Apply-ShapeStyle {
    param(
        $Shape,
        $Item,
        $Map = $null,
        [bool]$IsConnector = $false
    )

    $defaults = $Spec.defaults
    $lineColor = if ($Item.line_color) { $Item.line_color } elseif ($defaults.line_color) { $defaults.line_color } else { "#333333" }
    $textColor = if ($Item.text_color) { $Item.text_color } elseif ($defaults.text_color) { $defaults.text_color } else { "#111111" }
    $fillColor = if ($Item.fill_color) { $Item.fill_color } elseif ($defaults.fill_color) { $defaults.fill_color } else { "#FFFFFF" }
    $linePattern = if ($Item.line_pattern) { [string]$Item.line_pattern } elseif ($defaults.line_pattern) { [string]$defaults.line_pattern } else { "solid" }
    $fontSize = Resolve-RequestedFontSize -Item $Item -Defaults $defaults
    $fontFamily = Resolve-FontFamily -Item $Item -Defaults $defaults
    $lineWeight = if ($Item.line_weight) { [double]$Item.line_weight } else { 1.0 }
    $textAlign = if ($Item.text_align) { [string]$Item.text_align } else { "center" }
    $verticalAlign = if ($Item.vertical_align) { [string]$Item.vertical_align } elseif ($defaults.vertical_align) { [string]$defaults.vertical_align } else { "middle" }
    $fontBold = if ($null -ne $Item.font_bold) { [bool]$Item.font_bold } elseif ($null -ne $defaults.font_bold) { [bool]$defaults.font_bold } else { $false }
    $textPadding = if ($Item.text_padding) { $Item.text_padding } elseif ($defaults.text_padding) { $defaults.text_padding } else { $null }
    $resolvedFontSize = Resolve-FontSize -Item $Item -Defaults $defaults -Map $Map -RequestedFontSize $fontSize -IsConnector $IsConnector

    if (-not $IsConnector -and $PreserveColor) {
        Set-FormulaIfPresent -Shape $Shape -CellName "FillForegnd" -Formula (ConvertTo-RgbFormula -Color $fillColor -Fallback "#FFFFFF")
    }
    elseif (-not $IsConnector) {
        Set-FormulaIfPresent -Shape $Shape -CellName "FillForegnd" -Formula "RGB(255,255,255)"
    }

    Set-FormulaIfPresent -Shape $Shape -CellName "LineColor" -Formula (ConvertTo-RgbFormula -Color $lineColor -Fallback "#333333")
    Set-FormulaIfPresent -Shape $Shape -CellName "Char.Color" -Formula (ConvertTo-RgbFormula -Color $textColor -Fallback "#111111")
    if (-not [string]::IsNullOrWhiteSpace($fontFamily)) {
        $fontId = Resolve-FontId -FontFamily $fontFamily
        if ($null -ne $fontId) {
            Set-ResultIfPresent -Shape $Shape -CellName "Char.Font" -Value $fontId
        }
        else {
            Set-FormulaIfPresent -Shape $Shape -CellName "Char.Font" -Formula ('FONT("{0}")' -f $fontFamily.Replace('"', '""'))
        }
    }
    if ($fontBold) {
        Set-ResultIfPresent -Shape $Shape -CellName "Char.Style" -Value 1
    }
    Set-ResultIfPresent -Shape $Shape -CellName "LinePattern" -Value (Get-LinePatternCode -Pattern $linePattern)
    Set-ResultIfPresent -Shape $Shape -CellName "LineWeight" -Value ([Math]::Max(($lineWeight / 72.0), 0.01))
    Set-ResultIfPresent -Shape $Shape -CellName "Char.Size" -Value ([Math]::Max(($resolvedFontSize / 72.0), 0.08))

    switch ($textAlign.ToLowerInvariant()) {
        "left" { Set-ResultIfPresent -Shape $Shape -CellName "Para.HorzAlign" -Value 0 }
        "right" { Set-ResultIfPresent -Shape $Shape -CellName "Para.HorzAlign" -Value 2 }
        default { Set-ResultIfPresent -Shape $Shape -CellName "Para.HorzAlign" -Value 1 }
    }

    switch ($verticalAlign.ToLowerInvariant()) {
        "top" { Set-ResultIfPresent -Shape $Shape -CellName "VerticalAlign" -Value 0 }
        "bottom" { Set-ResultIfPresent -Shape $Shape -CellName "VerticalAlign" -Value 2 }
        default { Set-ResultIfPresent -Shape $Shape -CellName "VerticalAlign" -Value 1 }
    }

    if ($Map -and $textPadding) {
        if ($null -ne $textPadding.left) { Set-ResultIfPresent -Shape $Shape -CellName "LeftMargin" -Value (([double]$textPadding.left) * $Map.Scale) }
        if ($null -ne $textPadding.right) { Set-ResultIfPresent -Shape $Shape -CellName "RightMargin" -Value (([double]$textPadding.right) * $Map.Scale) }
        if ($null -ne $textPadding.top) { Set-ResultIfPresent -Shape $Shape -CellName "TopMargin" -Value (([double]$textPadding.top) * $Map.Scale) }
        if ($null -ne $textPadding.bottom) { Set-ResultIfPresent -Shape $Shape -CellName "BottomMargin" -Value (([double]$textPadding.bottom) * $Map.Scale) }
    }
}

function New-TextAnnotation {
    param($Page, $Item, $Map)

    $x = Convert-ToPageX -Map $Map -X ([double]$Item.x)
    $y = Convert-ToPageY -Map $Map -Y ([double]$Item.y)
    $w = if ($Item.width) { Convert-ToPageWidth -Map $Map -Width ([double]$Item.width) } else { 1.4 }
    $h = if ($Item.height) { Convert-ToPageHeight -Map $Map -Height ([double]$Item.height) } else { 0.3 }

    $shape = $Page.DrawRectangle($x - ($w / 2.0), $y - ($h / 2.0), $x + ($w / 2.0), $y + ($h / 2.0))
    $shape.Text = [string]$Item.text
    Apply-ShapeStyle -Shape $shape -Item $Item -Map $Map
    Set-ResultIfPresent -Shape $shape -CellName "FillPattern" -Value 0
    Set-ResultIfPresent -Shape $shape -CellName "LinePattern" -Value 0
    return $shape
}

function New-DiamondShape {
    param($Page, [double]$X, [double]$Y, [double]$Width, [double]$Height)

    $points = @(
        $X, ($Y + ($Height / 2.0)),
        ($X + ($Width / 2.0)), $Y,
        $X, ($Y - ($Height / 2.0)),
        ($X - ($Width / 2.0)), $Y,
        $X, ($Y + ($Height / 2.0))
    )

    return $Page.DrawPolyline(([double[]]$points), 0)
}

function New-BlockArrowRightShape {
    param($Page, [double]$X, [double]$Y, [double]$Width, [double]$Height)

    $halfW = $Width / 2.0
    $halfH = $Height / 2.0
    $shaftX = $X + ($halfW * 0.18)
    $shaftHalfH = $halfH * 0.42

    $points = @(
        ($X - $halfW), ($Y + $shaftHalfH),
        $shaftX, ($Y + $shaftHalfH),
        $shaftX, ($Y + $halfH),
        ($X + $halfW), $Y,
        $shaftX, ($Y - $halfH),
        $shaftX, ($Y - $shaftHalfH),
        ($X - $halfW), ($Y - $shaftHalfH),
        ($X - $halfW), ($Y + $shaftHalfH)
    )

    return $Page.DrawPolyline(([double[]]$points), 0)
}

function New-BlockArrowDownShape {
    param($Page, [double]$X, [double]$Y, [double]$Width, [double]$Height)

    $halfW = $Width / 2.0
    $halfH = $Height / 2.0
    $shaftHalfW = $halfW * 0.42
    $shaftTopY = $Y + ($halfH * 0.18)

    $points = @(
        ($X - $shaftHalfW), ($Y + $halfH),
        ($X + $shaftHalfW), ($Y + $halfH),
        ($X + $shaftHalfW), $shaftTopY,
        ($X + $halfW), $shaftTopY,
        $X, ($Y - $halfH),
        ($X - $halfW), $shaftTopY,
        ($X - $shaftHalfW), $shaftTopY,
        ($X - $shaftHalfW), ($Y + $halfH)
    )

    return $Page.DrawPolyline(([double[]]$points), 0)
}

function New-DiagramShape {
    param($Page, $Item, $Map)

    $x = Convert-ToPageX -Map $Map -X ([double]$Item.x)
    $y = Convert-ToPageY -Map $Map -Y ([double]$Item.y)
    $w = Convert-ToPageWidth -Map $Map -Width ([double]$Item.width)
    $h = Convert-ToPageHeight -Map $Map -Height ([double]$Item.height)

    $shapeType = if ($Item.shape) { [string]$Item.shape } else { "rectangle" }

    switch ($shapeType.ToLowerInvariant()) {
        "ellipse" { $shape = $Page.DrawOval($x - ($w / 2.0), $y - ($h / 2.0), $x + ($w / 2.0), $y + ($h / 2.0)) }
        "circle" { $diameter = [Math]::Min($w, $h); $shape = $Page.DrawOval($x - ($diameter / 2.0), $y - ($diameter / 2.0), $x + ($diameter / 2.0), $y + ($diameter / 2.0)) }
        "diamond" { $shape = New-DiamondShape -Page $Page -X $x -Y $y -Width $w -Height $h }
        "block-arrow-right" { $shape = New-BlockArrowRightShape -Page $Page -X $x -Y $y -Width $w -Height $h }
        "block-arrow-down" { $shape = New-BlockArrowDownShape -Page $Page -X $x -Y $y -Width $w -Height $h }
        default { $shape = $Page.DrawRectangle($x - ($w / 2.0), $y - ($h / 2.0), $x + ($w / 2.0), $y + ($h / 2.0)) }
    }

    if ($shapeType.ToLowerInvariant() -eq "rounded-rectangle") {
        Set-ResultIfPresent -Shape $shape -CellName "Rounding" -Value ([Math]::Min($w, $h) * 0.12)
    }
    elseif ($shapeType.ToLowerInvariant() -eq "terminator") {
        Set-ResultIfPresent -Shape $shape -CellName "Rounding" -Value ([Math]::Min($w, $h) * 0.25)
    }

    if ($Item.text) {
        $shape.Text = [string]$Item.text
    }

    if ($Item.rotation) {
        Set-ResultIfPresent -Shape $shape -CellName "Angle" -Value ([double]$Item.rotation * [Math]::PI / 180.0)
    }

    Apply-ShapeStyle -Shape $shape -Item $Item -Map $Map
    return $shape
}

function Get-PreparedImagePath {
    param(
        [string]$ImagePath,
        $Item
    )

    $defaults = if ($Spec.defaults) { $Spec.defaults } else { $null }
    $shouldTrim = $true
    if ($null -ne $Item.trim_white_margin) {
        $shouldTrim = [bool]$Item.trim_white_margin
    }
    elseif ($null -ne $defaults -and $null -ne $defaults.trim_white_margin) {
        $shouldTrim = [bool]$defaults.trim_white_margin
    }

    if (-not $shouldTrim) {
        return $ImagePath
    }

    $trimMode = if ($Item.trim_mode) { ([string]$Item.trim_mode).ToLowerInvariant() } elseif ($null -ne $defaults -and $defaults.trim_mode) { ([string]$defaults.trim_mode).ToLowerInvariant() } else { "all-content" }

    Add-Type -AssemblyName System.Drawing

    $bitmap = $null
    $croppedBitmap = $null
    try {
        $bitmap = New-Object System.Drawing.Bitmap($ImagePath)
        $contentMask = New-Object 'bool[]' ($bitmap.Width * $bitmap.Height)
        for ($y = 0; $y -lt $bitmap.Height; $y++) {
            for ($x = 0; $x -lt $bitmap.Width; $x++) {
                $pixel = $bitmap.GetPixel($x, $y)
                $isWhiteLike = ($pixel.A -lt 8) -or (($pixel.R -ge 248) -and ($pixel.G -ge 248) -and ($pixel.B -ge 248))
                if (-not $isWhiteLike) {
                    $contentMask[($y * $bitmap.Width) + $x] = $true
                }
            }
        }

        $minX = $bitmap.Width
        $minY = $bitmap.Height
        $maxX = -1
        $maxY = -1

        if ($trimMode -eq "largest-component") {
            $visited = New-Object 'bool[]' ($bitmap.Width * $bitmap.Height)
            $bestCount = 0
            for ($y = 0; $y -lt $bitmap.Height; $y++) {
                for ($x = 0; $x -lt $bitmap.Width; $x++) {
                    $idx = ($y * $bitmap.Width) + $x
                    if (-not $contentMask[$idx] -or $visited[$idx]) {
                        continue
                    }

                    $queue = New-Object System.Collections.Generic.Queue[System.Int32]
                    $queue.Enqueue($idx)
                    $visited[$idx] = $true

                    $compCount = 0
                    $compMinX = $bitmap.Width
                    $compMinY = $bitmap.Height
                    $compMaxX = -1
                    $compMaxY = -1

                    while ($queue.Count -gt 0) {
                        $current = $queue.Dequeue()
                        $cx = $current % $bitmap.Width
                        $cy = [int]($current / $bitmap.Width)
                        $compCount++
                        if ($cx -lt $compMinX) { $compMinX = $cx }
                        if ($cy -lt $compMinY) { $compMinY = $cy }
                        if ($cx -gt $compMaxX) { $compMaxX = $cx }
                        if ($cy -gt $compMaxY) { $compMaxY = $cy }

                        foreach ($offset in @(@(-1,0), @(1,0), @(0,-1), @(0,1))) {
                            $nx = $cx + $offset[0]
                            $ny = $cy + $offset[1]
                            if ($nx -lt 0 -or $ny -lt 0 -or $nx -ge $bitmap.Width -or $ny -ge $bitmap.Height) {
                                continue
                            }
                            $nIdx = ($ny * $bitmap.Width) + $nx
                            if ($contentMask[$nIdx] -and -not $visited[$nIdx]) {
                                $visited[$nIdx] = $true
                                $queue.Enqueue($nIdx)
                            }
                        }
                    }

                    if ($compCount -gt $bestCount) {
                        $bestCount = $compCount
                        $minX = $compMinX
                        $minY = $compMinY
                        $maxX = $compMaxX
                        $maxY = $compMaxY
                    }
                }
            }
        }
        else {
            for ($y = 0; $y -lt $bitmap.Height; $y++) {
                for ($x = 0; $x -lt $bitmap.Width; $x++) {
                    if ($contentMask[($y * $bitmap.Width) + $x]) {
                        if ($x -lt $minX) { $minX = $x }
                        if ($y -lt $minY) { $minY = $y }
                        if ($x -gt $maxX) { $maxX = $x }
                        if ($y -gt $maxY) { $maxY = $y }
                    }
                }
            }
        }

        if ($maxX -lt $minX -or $maxY -lt $minY) {
            return $ImagePath
        }

        $padding = 0
        if ($null -ne $Item.trim_padding) {
            $padding = [int][Math]::Max([double]$Item.trim_padding, 0.0)
        }
        elseif ($null -ne $defaults -and $null -ne $defaults.trim_padding) {
            $padding = [int][Math]::Max([double]$defaults.trim_padding, 0.0)
        }
        $cropX = [Math]::Max($minX - $padding, 0)
        $cropY = [Math]::Max($minY - $padding, 0)
        $cropWidth = [Math]::Min(($maxX - $minX + 1 + (2 * $padding)), ($bitmap.Width - $cropX))
        $cropHeight = [Math]::Min(($maxY - $minY + 1 + (2 * $padding)), ($bitmap.Height - $cropY))

        if ($cropX -eq 0 -and $cropY -eq 0 -and $cropWidth -eq $bitmap.Width -and $cropHeight -eq $bitmap.Height) {
            return $ImagePath
        }

        $rect = New-Object System.Drawing.Rectangle($cropX, $cropY, $cropWidth, $cropHeight)
        $croppedBitmap = $bitmap.Clone($rect, $bitmap.PixelFormat)
        $tempPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("codex-visio-img-" + [guid]::NewGuid().ToString("N") + ".png")
        $croppedBitmap.Save($tempPath, [System.Drawing.Imaging.ImageFormat]::Png)
        if ($script:PreparedImageTempPaths -eq $null) {
            $script:PreparedImageTempPaths = New-Object System.Collections.ArrayList
        }
        [void]$script:PreparedImageTempPaths.Add($tempPath)
        return $tempPath
    }
    finally {
        if ($null -ne $croppedBitmap) {
            $croppedBitmap.Dispose()
        }
        if ($null -ne $bitmap) {
            $bitmap.Dispose()
        }
    }
}

function New-ImageShape {
    param($Page, $Item, $Map)

    if (-not $Item.path) {
        throw "Image item is missing path."
    }

    $imagePath = [string]$Item.path
    if (-not (Test-Path -LiteralPath $imagePath)) {
        throw "Image file not found: $imagePath"
    }

    $trimmedPath = Get-PreparedImagePath -ImagePath $imagePath -Item $Item
    $shape = $Page.Import($trimmedPath)
    $x = Convert-ToPageX -Map $Map -X ([double]$Item.x)
    $y = Convert-ToPageY -Map $Map -Y ([double]$Item.y)
    $w = Convert-ToPageWidth -Map $Map -Width ([double]$Item.width)
    $h = Convert-ToPageHeight -Map $Map -Height ([double]$Item.height)

    Set-ResultIfPresent -Shape $shape -CellName "PinX" -Value $x
    Set-ResultIfPresent -Shape $shape -CellName "PinY" -Value $y
    Set-ResultIfPresent -Shape $shape -CellName "Width" -Value $w
    Set-ResultIfPresent -Shape $shape -CellName "Height" -Value $h

    if ($Item.send_to_back -eq $true) {
        try {
            $shape.SendToBack() | Out-Null
        }
        catch {
        }
    }
    elseif ($Item.bring_to_front -eq $true) {
        try {
            $shape.BringToFront() | Out-Null
        }
        catch {
        }
    }

    return $shape
}

function Get-GluePoint {
    param($Shape, [string]$Side)

    $normalized = if ([string]::IsNullOrWhiteSpace($Side)) { "auto" } else { $Side.ToLowerInvariant() }
    switch ($normalized) {
        "top" { return @(0.5, 1.0) }
        "right" { return @(1.0, 0.5) }
        "bottom" { return @(0.5, 0.0) }
        "left" { return @(0.0, 0.5) }
        default { return @(0.5, 0.5) }
    }
}

function Resolve-EdgeGluePoint {
    param(
        $Shape,
        $Edge,
        [string]$Endpoint
    )

    $glue = if ($Endpoint -eq "from") { $Edge.from_glue } else { $Edge.to_glue }
    if ($null -ne $glue -and $glue.Count -ge 2) {
        try {
            return @([double]$glue[0], [double]$glue[1])
        }
        catch {
        }
    }

    $side = if ($Endpoint -eq "from") { [string]$Edge.from_side } else { [string]$Edge.to_side }
    return Get-GluePoint -Shape $Shape -Side $side
}

function Get-SourceGluePoint {
    param(
        $Item,
        [double[]]$Glue
    )

    $itemRef = @($Item)[0]
    $glueRef = @($Glue)

    $x = [double](@($itemRef.x)[0])
    $y = [double](@($itemRef.y)[0])
    $width = [double](Get-ItemDimension -Item $itemRef -Dimension "width")
    $height = [double](Get-ItemDimension -Item $itemRef -Dimension "height")
    $glueX = [double](@($glueRef[0])[0])
    $glueY = [double](@($glueRef[1])[0])

    $left = [double]($x - ($width / 2.0))
    $top = [double]($y - ($height / 2.0))

    return [double[]]@(
        ([double]($left + ($width * $glueX))),
        ([double]($top + ($height * (1.0 - $glueY))))
    )
}

function New-PolylineConnectorShape {
    param(
        $Page,
        $Map,
        $Edge,
        $FromItem,
        $ToItem
    )

    $fromGlue = Resolve-EdgeGluePoint -Shape $null -Edge $Edge -Endpoint "from"
    $toGlue = Resolve-EdgeGluePoint -Shape $null -Edge $Edge -Endpoint "to"
    $fromPoint = Get-SourceGluePoint -Item $FromItem -Glue $fromGlue
    $toPoint = Get-SourceGluePoint -Item $ToItem -Glue $toGlue

    $points = New-Object System.Collections.Generic.List[double]
    $startPoint = Convert-PointToPage -Map $Map -X $fromPoint[0] -Y $fromPoint[1]
    [void]$points.Add([double]$startPoint[0])
    [void]$points.Add([double]$startPoint[1])

    foreach ($waypoint in @($Edge.waypoints)) {
        if ($waypoint.Count -lt 2) {
            continue
        }
        $pagePoint = Convert-PointToPage -Map $Map -X ([double]$waypoint[0]) -Y ([double]$waypoint[1])
        [void]$points.Add([double]$pagePoint[0])
        [void]$points.Add([double]$pagePoint[1])
    }

    $endPoint = Convert-PointToPage -Map $Map -X $toPoint[0] -Y $toPoint[1]
    [void]$points.Add([double]$endPoint[0])
    [void]$points.Add([double]$endPoint[1])

    $shape = $Page.DrawPolyline(($points.ToArray()), 0)
    Apply-ShapeStyle -Shape $shape -Item $Edge -Map $Map -IsConnector $true
    Set-ResultIfPresent -Shape $shape -CellName "EndArrow" -Value (Get-EndArrowCode -Arrow ([string]$Edge.arrow_end))
    $pathLength = Get-ConnectorPathLength -Points ($points.ToArray())
    $arrowSize = if ($Edge.arrow_size) { [string]$Edge.arrow_size } elseif ($Spec.defaults.arrow_size) { [string]$Spec.defaults.arrow_size } else { Resolve-AutoArrowSize -Length $pathLength }
    Set-ResultIfPresent -Shape $shape -CellName "EndArrowSize" -Value (Get-ArrowSizeCode -ArrowSize $arrowSize)
    Set-ResultIfPresent -Shape $shape -CellName "FillPattern" -Value 0
    if ($Edge.text) {
        $shape.Text = [string]$Edge.text
    }

    return $shape
}

function Get-ConnectorPathLength {
    param([double[]]$Points)

    if (-not $Points -or $Points.Count -lt 4) {
        return 0.0
    }

    $length = 0.0
    for ($i = 0; $i -lt ($Points.Count - 2); $i += 2) {
        $x1 = [double]$Points[$i]
        $y1 = [double]$Points[$i + 1]
        $x2 = [double]$Points[$i + 2]
        $y2 = [double]$Points[$i + 3]
        $length += [Math]::Sqrt((($x2 - $x1) * ($x2 - $x1)) + (($y2 - $y1) * ($y2 - $y1)))
    }

    return $length
}

function Resolve-AutoArrowSize {
    param([double]$Length)

    if ($Length -lt 0.35) { return "tiny" }
    if ($Length -lt 0.7) { return "small" }
    if ($Length -lt 1.25) { return "medium" }
    return "large"
}

function Get-LayoutItems {
    param($Spec)

    $items = @()
    if ($Spec.containers) { $items += @($Spec.containers) }
    if ($Spec.nodes) { $items += @($Spec.nodes) }
    if ($Spec.annotations) { $items += @($Spec.annotations) }
    if ($Spec.images) { $items += @($Spec.images) }
    return $items
}

function Get-LayoutIndex {
    param($Spec)

    $index = @{}
    foreach ($item in @(Get-LayoutItems -Spec $Spec)) {
        if ($item.id) {
            $index[[string]$item.id] = $item
        }
    }

    return $index
}

function Get-ItemDimension {
    param($Item, [string]$Dimension)

    $itemRef = @($Item)[0]

    switch ($Dimension.ToLowerInvariant()) {
        "width" {
            if ($null -ne $itemRef.width) { return [double](@($itemRef.width)[0]) }
            return 0.0
        }
        "height" {
            if ($null -ne $itemRef.height) { return [double](@($itemRef.height)[0]) }
            return 0.0
        }
        default { return 0.0 }
    }
}

function Set-ItemDimension {
    param($Item, [string]$Dimension, [double]$Value)

    switch ($Dimension.ToLowerInvariant()) {
        "width" { $Item.width = $Value }
        "height" { $Item.height = $Value }
    }
}

function Get-ItemAnchorValue {
    param(
        $Item,
        [string]$Axis,
        [string]$Anchor = "center"
    )

    $axisName = $Axis.ToLowerInvariant()
    $anchorName = $Anchor.ToLowerInvariant()

    if ($axisName -eq "x") {
        $x = [double]$Item.x
        $w = Get-ItemDimension -Item $Item -Dimension "width"
        switch ($anchorName) {
            "left" { return $x - ($w / 2.0) }
            "right" { return $x + ($w / 2.0) }
            default { return $x }
        }
    }

    $y = [double]$Item.y
    $h = Get-ItemDimension -Item $Item -Dimension "height"
    switch ($anchorName) {
        "top" { return $y - ($h / 2.0) }
        "bottom" { return $y + ($h / 2.0) }
        default { return $y }
    }
}

function Set-ItemAnchorValue {
    param(
        $Item,
        [string]$Axis,
        [string]$Anchor,
        [double]$Value
    )

    $axisName = $Axis.ToLowerInvariant()
    $anchorName = $Anchor.ToLowerInvariant()

    if ($axisName -eq "x") {
        $w = Get-ItemDimension -Item $Item -Dimension "width"
        switch ($anchorName) {
            "left" { $Item.x = $Value + ($w / 2.0) }
            "right" { $Item.x = $Value - ($w / 2.0) }
            default { $Item.x = $Value }
        }
        return
    }

    $h = Get-ItemDimension -Item $Item -Dimension "height"
    switch ($anchorName) {
        "top" { $Item.y = $Value + ($h / 2.0) }
        "bottom" { $Item.y = $Value - ($h / 2.0) }
        default { $Item.y = $Value }
    }
}

function Resolve-LayoutTargetValue {
    param(
        $Rule,
        $ItemIndex,
        [string]$Axis,
        [string]$Anchor,
        $FallbackItem = $null
    )

    if ($null -ne $Rule.value) {
        return [double]$Rule.value
    }

    if ($Rule.reference_id -and $ItemIndex.ContainsKey([string]$Rule.reference_id)) {
        $referenceItem = $ItemIndex[[string]$Rule.reference_id]
        $referenceAnchor = if ($Rule.reference_anchor) { [string]$Rule.reference_anchor } else { $Anchor }
        return Get-ItemAnchorValue -Item $referenceItem -Axis $Axis -Anchor $referenceAnchor
    }

    if ($null -ne $FallbackItem) {
        return Get-ItemAnchorValue -Item $FallbackItem -Axis $Axis -Anchor $Anchor
    }

    return $null
}

function Apply-LayoutConstraints {
    param($Spec)

    $rules = @($Spec.layout)
    if ($rules.Count -eq 0) {
        return
    }

    $itemIndex = Get-LayoutIndex -Spec $Spec

    foreach ($rule in $rules) {
        if (-not $rule.type) {
            continue
        }

        $type = [string]$rule.type
        switch ($type.ToLowerInvariant()) {
            "align" {
                $ids = @($rule.ids)
                if ($ids.Count -eq 0) { continue }
                $items = @($ids | ForEach-Object { if ($itemIndex.ContainsKey([string]$_)) { $itemIndex[[string]$_] } }) | Where-Object { $null -ne $_ }
                if ($items.Count -eq 0) { continue }
                $axis = if ($rule.axis) { [string]$rule.axis } else { "x" }
                $anchor = if ($rule.anchor) { [string]$rule.anchor } else { "center" }
                $targetValue = Resolve-LayoutTargetValue -Rule $rule -ItemIndex $itemIndex -Axis $axis -Anchor $anchor -FallbackItem $items[0]
                if ($null -eq $targetValue) { continue }
                foreach ($item in $items) {
                    Set-ItemAnchorValue -Item $item -Axis $axis -Anchor $anchor -Value ([double]$targetValue)
                }
            }
            "same_size" {
                $ids = @($rule.ids)
                if ($ids.Count -eq 0) { continue }
                $items = @($ids | ForEach-Object { if ($itemIndex.ContainsKey([string]$_)) { $itemIndex[[string]$_] } }) | Where-Object { $null -ne $_ }
                if ($items.Count -eq 0) { continue }
                $dimension = if ($rule.dimension) { [string]$rule.dimension } else { "width" }
                $dimensions = switch ($dimension.ToLowerInvariant()) {
                    "both" { @("width", "height") }
                    default { @($dimension.ToLowerInvariant()) }
                }
                $referenceItem = $null
                if ($rule.reference_id -and $itemIndex.ContainsKey([string]$rule.reference_id)) {
                    $referenceItem = $itemIndex[[string]$rule.reference_id]
                }
                if ($null -eq $referenceItem) {
                    $referenceItem = $items[0]
                }
                foreach ($dim in $dimensions) {
                    $explicitProperty = $rule.PSObject.Properties[$dim]
                    $targetValue = if ($null -ne $explicitProperty) { [double]$explicitProperty.Value } else { Get-ItemDimension -Item $referenceItem -Dimension $dim }
                    foreach ($item in $items) {
                        Set-ItemDimension -Item $item -Dimension $dim -Value $targetValue
                    }
                }
            }
            "distribute" {
                $ids = @($rule.ids)
                if ($ids.Count -lt 2) { continue }
                $items = @($ids | ForEach-Object { if ($itemIndex.ContainsKey([string]$_)) { $itemIndex[[string]$_] } }) | Where-Object { $null -ne $_ }
                if ($items.Count -lt 2) { continue }
                $axis = if ($rule.axis) { [string]$rule.axis } else { "y" }
                $anchor = if ($rule.anchor) { [string]$rule.anchor } else { "center" }
                $startValue = if ($null -ne $rule.start) { [double]$rule.start } else { Get-ItemAnchorValue -Item $items[0] -Axis $axis -Anchor $anchor }
                if ($null -ne $rule.gap) {
                    $gapValue = [double]$rule.gap
                }
                elseif ($null -ne $rule.end) {
                    $gapValue = ([double]$rule.end - $startValue) / [Math]::Max(($items.Count - 1), 1)
                }
                else {
                    $endValue = Get-ItemAnchorValue -Item $items[$items.Count - 1] -Axis $axis -Anchor $anchor
                    $gapValue = ($endValue - $startValue) / [Math]::Max(($items.Count - 1), 1)
                }
                for ($i = 0; $i -lt $items.Count; $i++) {
                    Set-ItemAnchorValue -Item $items[$i] -Axis $axis -Anchor $anchor -Value ($startValue + ($i * $gapValue))
                }
            }
            "offset" {
                if (-not $rule.id -or -not $rule.reference_id) { continue }
                $itemId = [string]$rule.id
                $referenceId = [string]$rule.reference_id
                if (-not $itemIndex.ContainsKey($itemId) -or -not $itemIndex.ContainsKey($referenceId)) { continue }
                $item = $itemIndex[$itemId]
                $referenceItem = $itemIndex[$referenceId]
                $dx = if ($null -ne $rule.dx) { [double]$rule.dx } else { 0.0 }
                $dy = if ($null -ne $rule.dy) { [double]$rule.dy } else { 0.0 }
                $item.x = [double]$referenceItem.x + $dx
                $item.y = [double]$referenceItem.y + $dy
            }
        }
    }
}

function Resolve-SlotItems {
    param($Spec)

    $itemIndex = Get-LayoutIndex -Spec $Spec

    foreach ($item in @(Get-LayoutItems -Spec $Spec)) {
        if (-not $item.slot) {
            continue
        }

        $slot = $item.slot
        if (-not $slot.reference_id) {
            continue
        }

        $referenceId = [string]$slot.reference_id
        if (-not $itemIndex.ContainsKey($referenceId)) {
            continue
        }

        $hostItem = $itemIndex[$referenceId]
        $hostX = [double]$hostItem.x
        $hostY = [double]$hostItem.y
        $hostW = Get-ItemDimension -Item $hostItem -Dimension "width"
        $hostH = Get-ItemDimension -Item $hostItem -Dimension "height"
        $itemW = Get-ItemDimension -Item $item -Dimension "width"
        $itemH = Get-ItemDimension -Item $item -Dimension "height"
        $dx = if ($null -ne $slot.dx) { [double]$slot.dx } else { 0.0 }
        $dy = if ($null -ne $slot.dy) { [double]$slot.dy } else { 0.0 }
        $side = if ($slot.side) { [string]$slot.side } else { "left" }
        $align = if ($slot.align) { [string]$slot.align } else { "center" }

        switch ($side.ToLowerInvariant()) {
            "left" { $item.x = $hostX - ($hostW / 2.0) + ($itemW / 2.0) + $dx }
            "right" { $item.x = $hostX + ($hostW / 2.0) - ($itemW / 2.0) + $dx }
            "top" { $item.y = $hostY - ($hostH / 2.0) + ($itemH / 2.0) + $dy }
            "bottom" { $item.y = $hostY + ($hostH / 2.0) - ($itemH / 2.0) + $dy }
            "center" {
                $item.x = $hostX + $dx
                $item.y = $hostY + $dy
            }
        }

        if ($side.ToLowerInvariant() -in @("left", "right", "center")) {
            switch ($align.ToLowerInvariant()) {
                "start" { $item.y = $hostY - ($hostH / 2.0) + ($itemH / 2.0) + $dy }
                "end" { $item.y = $hostY + ($hostH / 2.0) - ($itemH / 2.0) + $dy }
                default { $item.y = $hostY + $dy }
            }
        }

        if ($side.ToLowerInvariant() -in @("top", "bottom", "center")) {
            switch ($align.ToLowerInvariant()) {
                "start" { $item.x = $hostX - ($hostW / 2.0) + ($itemW / 2.0) + $dx }
                "end" { $item.x = $hostX + ($hostW / 2.0) - ($itemW / 2.0) + $dx }
                default { $item.x = $hostX + $dx }
            }
        }
    }
}

function Resolve-VerticalOverlaps {
    param($Spec)

    $candidates = @(
        @($Spec.nodes) +
        @($Spec.annotations) |
            Where-Object {
                $_.id -and $_.shape -and $_.shape -ne "none" -and $null -ne $_.width -and $null -ne $_.height -and ([double]$_.width -gt 0) -and ([double]$_.height -gt 0)
            }
    )

    if ($candidates.Count -lt 2) {
        return
    }

    $groups = New-Object System.Collections.Generic.List[object]
    foreach ($item in ($candidates | Sort-Object { [double]$_.x })) {
        $centerX = [double]$item.x
        $width = [double]$item.width
        $placed = $false
        foreach ($group in $groups) {
            $threshold = [Math]::Max(20.0, [Math]::Min([double]$group.avgWidth, $width) * 0.25)
            if ([Math]::Abs($group.centerX - $centerX) -le $threshold) {
                $prevCount = [double]$group.items.Count
                $group.items.Add($item) | Out-Null
                $group.centerX = (($group.centerX * $prevCount) + $centerX) / ($prevCount + 1.0)
                $group.avgWidth = (($group.avgWidth * $prevCount) + $width) / ($prevCount + 1.0)
                $placed = $true
                break
            }
        }

        if (-not $placed) {
            $bucket = [pscustomobject]@{
                centerX = $centerX
                avgWidth = $width
                items = New-Object System.Collections.ArrayList
            }
            [void]$bucket.items.Add($item)
            [void]$groups.Add($bucket)
        }
    }

    foreach ($group in $groups) {
        $ordered = @($group.items | Sort-Object { [double]$_.y })
        if ($ordered.Count -lt 2) { continue }
        $cursorBottom = [double]$ordered[0].y + ([double]$ordered[0].height / 2.0)
        for ($i = 1; $i -lt $ordered.Count; $i++) {
            $item = $ordered[$i]
            $halfH = [double]$item.height / 2.0
            $minTop = $cursorBottom + 4.0
            $currentTop = [double]$item.y - $halfH
            if ($currentTop -lt $minTop) {
                $item.y = $minTop + $halfH
            }
            $cursorBottom = [double]$item.y + $halfH
        }
    }
}

if (-not (Test-Path -LiteralPath $SpecPath)) {
    throw "Spec file not found: $SpecPath"
}

$specText = [System.IO.File]::ReadAllText($SpecPath, [System.Text.Encoding]::UTF8)
$Spec = $specText | ConvertFrom-Json

if (-not $Spec.canvas -or -not $Spec.canvas.width -or -not $Spec.canvas.height) {
    throw "Spec must define canvas.width and canvas.height."
}

Apply-LayoutConstraints -Spec $Spec
Resolve-SlotItems -Spec $Spec
Resolve-VerticalOverlaps -Spec $Spec

$resolvedOrientation = Resolve-Orientation -Spec $Spec
$pageDims = Resolve-PageDimensions -ResolvedOrientation $resolvedOrientation
$pageMargin = Get-PageMargin -Spec $Spec
$map = New-CoordinateMapper -CanvasWidth ([double]$Spec.canvas.width) -CanvasHeight ([double]$Spec.canvas.height) -PageWidth $pageDims.Width -PageHeight $pageDims.Height -Margin $pageMargin

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDir = Split-Path -Path $outputFullPath -Parent
if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}
Test-OutputFileWritable -Path $outputFullPath

$visio = $null
$document = $null
$page = $null
$createdComObjects = New-Object System.Collections.ArrayList
$renderedFrontEntries = New-Object System.Collections.ArrayList
$previousAlertResponse = 0

try {
$visio = New-Object -ComObject Visio.Application
$script:VisioApp = $visio
$visio.Visible = $Visible
$previousAlertResponse = $visio.AlertResponse
$visio.AlertResponse = 7
$document = $visio.Documents.Add("")
$script:VisioDocument = $document
$page = $visio.ActivePage

    Set-ResultIfPresent -Shape $page.PageSheet -CellName "PageWidth" -Value $pageDims.Width
    Set-ResultIfPresent -Shape $page.PageSheet -CellName "PageHeight" -Value $pageDims.Height

    if ($Spec.title) {
        $page.Name = [string]$Spec.title
    }

    if ($Spec.page.background) {
        Set-FormulaIfPresent -Shape $page.PageSheet -CellName "PageColor" -Formula (ConvertTo-RgbFormula -Color ([string]$Spec.page.background) -Fallback "#FFFFFF")
    }

    $shapeIndex = @{}
    $specItemIndex = Get-LayoutIndex -Spec $Spec

    $renderQueue = New-Object System.Collections.ArrayList
    foreach ($containerItem in @($Spec.containers)) {
        [void]$renderQueue.Add([pscustomobject]@{ kind = "shape"; item = $containerItem; z = $(if ($containerItem.z_index) { [int]$containerItem.z_index } else { 50 }) })
    }
    foreach ($nodeItem in @($Spec.nodes)) {
        [void]$renderQueue.Add([pscustomobject]@{ kind = "shape"; item = $nodeItem; z = $(if ($nodeItem.z_index) { [int]$nodeItem.z_index } else { 50 }) })
    }
    foreach ($annotationItem in @($Spec.annotations)) {
        [void]$renderQueue.Add([pscustomobject]@{ kind = "shape"; item = $annotationItem; z = $(if ($annotationItem.z_index) { [int]$annotationItem.z_index } else { 50 }) })
    }
    foreach ($imageItem in @($Spec.images)) {
        [void]$renderQueue.Add([pscustomobject]@{ kind = "image"; item = $imageItem; z = $(if ($imageItem.z_index) { [int]$imageItem.z_index } else { 60 }) })
    }

    $orderedItems = $renderQueue | Sort-Object z

    foreach ($renderEntry in $orderedItems) {
        $item = $renderEntry.item

        if ($renderEntry.kind -eq "image") {
            $imageShape = New-ImageShape -Page $page -Item $item -Map $map
            [void]$createdComObjects.Add($imageShape)
            if ($item.send_to_back -ne $true) {
                [void]$renderedFrontEntries.Add([pscustomobject]@{ shape = $imageShape; z = $renderEntry.z })
            }
            continue
        }

        $shape = if ($item.shape -and [string]$item.shape -eq "none") {
            New-TextAnnotation -Page $page -Item $item -Map $map
        }
        elseif (-not $item.width -or -not $item.height) {
            New-TextAnnotation -Page $page -Item $item -Map $map
        }
        else {
            New-DiagramShape -Page $page -Item $item -Map $map
        }

        if ($item.id) {
            $shapeIndex[[string]$item.id] = $shape
        }

        [void]$createdComObjects.Add($shape)

        if ($item.send_to_back -eq $true) {
            try {
                $shape.SendToBack() | Out-Null
            }
            catch {
            }
        }
        else {
            [void]$renderedFrontEntries.Add([pscustomobject]@{ shape = $shape; z = $renderEntry.z })
        }
    }

    foreach ($edge in @($Spec.connectors)) {
        if (-not $edge.from -or -not $edge.to) {
            continue
        }

        $fromId = [string]$edge.from
        $toId = [string]$edge.to
        if (-not $shapeIndex.ContainsKey($fromId) -or -not $shapeIndex.ContainsKey($toId)) {
            continue
        }

        if ($edge.uncertain -eq $true) {
            $edge.line_color = $UncertainConnectorColor
        }

        $renderMode = if ($edge.render_mode) { ([string]$edge.render_mode).ToLowerInvariant() } elseif ($edge.waypoints) { "polyline" } else { "dynamic" }

        if ($renderMode -eq "polyline" -and $specItemIndex.ContainsKey($fromId) -and $specItemIndex.ContainsKey($toId)) {
            $connector = New-PolylineConnectorShape -Page $page -Map $map -Edge $edge -FromItem $specItemIndex[$fromId] -ToItem $specItemIndex[$toId]
        }
        else {
            $connectorMaster = $visio.Application.ConnectorToolDataObject
            $connector = $page.Drop($connectorMaster, 0, 0)

            $fromShape = $shapeIndex[$fromId]
            $toShape = $shapeIndex[$toId]
            $fromPos = Resolve-EdgeGluePoint -Shape $fromShape -Edge $edge -Endpoint "from"
            $toPos = Resolve-EdgeGluePoint -Shape $toShape -Edge $edge -Endpoint "to"

            $connector.CellsU("BeginX").GlueToPos($fromShape, $fromPos[0], $fromPos[1])
            $connector.CellsU("EndX").GlueToPos($toShape, $toPos[0], $toPos[1])

            if ($edge.text) {
                $connector.Text = [string]$edge.text
            }

            Apply-ShapeStyle -Shape $connector -Item $edge -Map $map -IsConnector $true
            Set-ResultIfPresent -Shape $connector -CellName "EndArrow" -Value (Get-EndArrowCode -Arrow ([string]$edge.arrow_end))
            $connectorLength = Get-ConnectorPathLength -Points @($fromPos[0], $fromPos[1], $toPos[0], $toPos[1])
            $arrowSize = if ($edge.arrow_size) { [string]$edge.arrow_size } elseif ($Spec.defaults.arrow_size) { [string]$Spec.defaults.arrow_size } else { Resolve-AutoArrowSize -Length $connectorLength }
            Set-ResultIfPresent -Shape $connector -CellName "EndArrowSize" -Value (Get-ArrowSizeCode -ArrowSize $arrowSize)
        }

        [void]$createdComObjects.Add($connector)
    }

    foreach ($frontEntry in @($renderedFrontEntries | Sort-Object z)) {
        try {
            $frontEntry.shape.BringToFront() | Out-Null
        }
        catch {
        }
    }

    $document.SaveAs($outputFullPath) | Out-Null
    $document.Saved = $true

    [pscustomobject]@{
        output = $outputFullPath
        title = $Spec.title
        nodes = @($Spec.nodes).Count
        containers = @($Spec.containers).Count
        connectors = @($Spec.connectors).Count
        uncertain_connectors = @($Spec.connectors | Where-Object { $_.uncertain -eq $true }).Count
    } | ConvertTo-Json -Compress
}
finally {
    $shapeIndex = $null
    $orderedItems = $null
    $renderQueue = $null
    $renderedFrontEntries = $null

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if ($null -ne $document) {
        try {
            $document.Close() | Out-Null
        }
        catch {
        }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
    }

    if ($null -ne $visio) {
        try {
            $visio.Quit() | Out-Null
        }
        catch {
        }
    }

    foreach ($tempPath in @($script:PreparedImageTempPaths)) {
        try {
            if (Test-Path -LiteralPath $tempPath) {
                Remove-Item -LiteralPath $tempPath -Force
            }
        }
        catch {
        }
    }
    $script:PreparedImageTempPaths = $null

    for ($i = $createdComObjects.Count - 1; $i -ge 0; $i--) {
        $comObject = $createdComObjects[$i]
        if ($null -ne $comObject) {
            try {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) | Out-Null
            }
            catch {
            }
        }
    }

    if ($null -ne $page) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($page) | Out-Null
        }
        catch {
        }
    }

    if ($null -ne $document) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
        }
        catch {
        }
    }

    if ($null -ne $visio) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($visio) | Out-Null
        }
        catch {
        }
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
