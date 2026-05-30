[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-Result {
    param(
        [bool]$Installed,
        [string]$Name,
        [string]$Version,
        [string]$Message
    )

    [pscustomobject]@{
        installed = $Installed
        name = $Name
        version = $Version
        message = $Message
    }
}

$visio = $null

try {
    $type = [type]::GetTypeFromProgID("Visio.Application")
    if (-not $type) {
        New-Result -Installed $false -Name "" -Version "" -Message "Visio COM ProgID not found." |
            ConvertTo-Json -Compress
        exit 1
    }

    $visio = New-Object -ComObject Visio.Application
    $visio.Visible = $false

    New-Result -Installed $true -Name $visio.Name -Version $visio.Version -Message "Visio COM available." |
        ConvertTo-Json -Compress
}
catch {
    New-Result -Installed $false -Name "" -Version "" -Message $_.Exception.Message |
        ConvertTo-Json -Compress
    exit 1
}
finally {
    if ($null -ne $visio) {
        try {
            $visio.Quit()
        }
        catch {
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($visio) | Out-Null
    }
}
