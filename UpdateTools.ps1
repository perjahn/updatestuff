Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

[string] $starttime = ([DateTime]::UtcNow).ToString("yyyy-MM")
[string] $logfile = "UpdateTools_$($starttime).log"
Start-Transcript "UpdateTools_transcript_$($starttime).log" -Append

function Main() {
    $watch = [Diagnostics.Stopwatch]::StartNew()

    UpdateChocoTools

    UpdatePSModules

    .\UpdateGit.ps1
    .\UpdateVS.ps1

    Log "Done: $($watch.Elapsed)" Cyan
}

function UpdateChocoTools() {
    Log "Updating Chocolatey tools..."
    choco update all -y
}

function UpdatePSModules() {
    Log "Updating Powershell modules..."
    Install-Module -Name Az -AllowClobber -Scope AllUsers -Force
}

function Log([string] $message, $color) {
    ([DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + ": " + $message) | Add-Content $logfile

    if ($color) {
        Write-Host $message -f $color
    }
    else {
        Write-Host $message -f Green
    }
}

Main
