[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ImagePath,

    [string]$SpecPath = "",

    [string]$OutputPath = "",

    [ValidateSet("Auto", "Landscape", "Portrait")]
    [string]$Orientation = "Auto",

    [ValidateSet("Auto", "A4", "A3", "Letter", "Legal")]
    [string]$PageSize = "Auto",

    [bool]$PreserveColor = $true,

    [string]$UncertainConnectorColor = "#FF00FF",

    [switch]$Visible,

    [switch]$CleanupIntermediate,

    [switch]$SkipEnvironmentTest
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$generatorPath = Join-Path $scriptDir "Convert-DiagramSpecToVisio.ps1"
$probePath = Join-Path $scriptDir "Test-VisioEnvironment.ps1"

if (-not (Test-Path -LiteralPath $ImagePath)) {
    throw "Image file not found: $ImagePath"
}

$imageFullPath = [System.IO.Path]::GetFullPath($ImagePath)
$extension = [System.IO.Path]::GetExtension($imageFullPath).ToLowerInvariant()
if ($extension -notin @(".png", ".jpg", ".jpeg", ".bmp")) {
    throw "Unsupported image extension: $extension. Use PNG, JPG, JPEG, or BMP."
}

if ([string]::IsNullOrWhiteSpace($SpecPath)) {
    $SpecPath = [System.IO.Path]::ChangeExtension($imageFullPath, ".json")
}

if (-not (Test-Path -LiteralPath $SpecPath)) {
    throw "Spec file not found: $SpecPath. First extract the diagram into JSON using references\diagram-spec.md, then rerun this wrapper."
}

$specFullPath = [System.IO.Path]::GetFullPath($SpecPath)

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = [System.IO.Path]::ChangeExtension($imageFullPath, ".vsdx")
}

$outputFullPath = [System.IO.Path]::GetFullPath($OutputPath)

if (-not $SkipEnvironmentTest) {
    & $probePath | Out-Null
}

$result = & $generatorPath `
    -SpecPath $specFullPath `
    -OutputPath $outputFullPath `
    -Orientation $Orientation `
    -PageSize $PageSize `
    -PreserveColor ([bool]$PreserveColor) `
    -UncertainConnectorColor $UncertainConnectorColor `
    -Visible ([bool]$Visible.IsPresent)

if ($CleanupIntermediate -and (Test-Path -LiteralPath $outputFullPath)) {
    if ([System.IO.Path]::GetExtension($specFullPath).ToLowerInvariant() -eq ".json") {
        Remove-Item -LiteralPath $specFullPath -Force
    }
}

$result
