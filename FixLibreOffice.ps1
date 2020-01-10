Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

# Remove software with cves from libreoffice.

[string] $starttime = ([DateTime]::UtcNow).ToString("yyyy-MM")
[string] $logfile = "FixLibreOffice_$($starttime).log"
Start-Transcript ("FixLibreOffice_transcript_$($starttime).log") -Append

function Main() {
    [string] $path = "C:\Program Files\Java"
    if (Test-Path $path) {
        Log "Removing: '$path'"
        rd -recurse -force $path
    }

    [string] $path = "C:\Program Files\LibreOffice\program\python-core-3.5.9"
    if (Test-Path $path) {
        Log "Removing: '$path'"
        rd -recurse -force $path
    }

    [string] $path = "C:\Program Files\LibreOffice\program\python.exe"
    if (Test-Path $path) {
        Log "Removing: '$path'"
        del $path
    }

    $files = dir -r "C:\Program Files\LibreOffice\python.exe"
    $files | % {
        [string] $filename = $_.FullName
        Log "Found python: '$filename'"
    }
}

function Log([string] $message, $color) {
    ([DateTime]::UtcNow.ToString("yyyy:MM:dd HH-mm-ss") + ": " + $message) | Add-Content $logfile

    if ($color) {
        Write-Host $message -f $color
    }
    else {
        Write-Host $message -f Green
    }
}

Main
