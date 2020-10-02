Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

[string] $transcriptfilePre = Join-Path (pwd).Path "pre_transcript.log"
[string] $logfilePre = Join-Path (pwd).Path "pre_action.log"

Start-Transcript -Path $transcriptfilePre -Append

function Main() {
    Import-Module "IISAdministration"

    [string] $serviceName = "OctopusDeploy"

    Log "Stopping service: '$serviceName'"
    Stop-Service $serviceName

    Start-Sleep 10

    [string] $poolName = "DefaultAppPool"
    [string] $siteName = "Default Web Site"

    Log "Starting pool: '$poolName'"
    $pool = Get-IISAppPool $poolName
    $pool.Start()

    Log "Starting website: '$siteName'"
    Start-IISSite $siteName
}

function Log([string] $message, $color) {
    ([DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + ": " + $message) | Add-Content $logfilePre

    if ($color) {
        Write-Host $message -f $color
    }
    else {
        Write-Host $message -f Green
    }
}

Main
