Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

[string] $transcriptfilePost = Join-Path (pwd).Path "post_transcript.log"
[string] $logfilePost = Join-Path (pwd).Path "post_action.log"

Start-Transcript -Path $transcriptfilePost -Append

function Main() {
    Import-Module "IISAdministration"

    [string] $poolName = "DefaultAppPool"
    [string] $siteName = "Default Web Site"

    Log "Stopping website: '$siteName'"
    Stop-IISSite $siteName -confirm:$false

    Log "Stopping pool: '$poolName'"
    $pool = Get-IISAppPool $poolName
    $pool.Stop()

    Start-Sleep 10

    [string] $serviceName = "OctopusDeploy"

    Log "Starting service: '$serviceName'"
    Start-Service $serviceName
}

function Log([string] $message, $color) {
    ([DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + ": " + $message) | Add-Content $logfilePost

    if ($color) {
        Write-Host $message -f $color
    }
    else {
        Write-Host $message -f Green
    }
}

Main
