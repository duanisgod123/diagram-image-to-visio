[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$VsdxPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [int]$PageNumber = 1,

    [switch]$Visible
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $VsdxPath)) {
    throw "VSDX file not found: $VsdxPath"
}

$outDir = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$visio = $null
$doc = $null
$previousAlertResponse = $null

try {
    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = [bool]$Visible.IsPresent
    $previousAlertResponse = $visio.AlertResponse
    $visio.AlertResponse = 7

    $doc = $visio.Documents.Open($VsdxPath)
    if ($PageNumber -lt 1 -or $PageNumber -gt $doc.Pages.Count) {
        throw "PageNumber $PageNumber out of range. Document has $($doc.Pages.Count) page(s)."
    }

    $page = $doc.Pages.Item($PageNumber)
    $page.Export($OutputPath)
    [pscustomobject]@{
        output = [System.IO.Path]::GetFullPath($OutputPath)
        source = [System.IO.Path]::GetFullPath($VsdxPath)
        page = $PageNumber
    } | ConvertTo-Json -Compress
}
finally {
    if ($null -ne $doc) {
        try { $doc.Close() } catch {}
    }
    if ($null -ne $visio) {
        if ($null -ne $previousAlertResponse) {
            try { $visio.AlertResponse = $previousAlertResponse } catch {}
        }
        try { $visio.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($visio) | Out-Null
    }
}
