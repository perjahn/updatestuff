Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

function UpdateVS() {
    Log "Updating Visual Studio..."

    $watch = [Diagnostics.Stopwatch]::StartNew()

    [string] $filename = Download-VisualStudio
    if (!$filename) {
        return
    }

    Install-VisualStudio

    Log "Done: $($watch.Elapsed)"
}

function Download-VisualStudio() {
    [string] $filename = "vs_enterprise.exe"
    del $filename -ErrorAction SilentlyContinue

    if (Test-Path $filename) {
        [string] $oldfilename = "vs_enterprise_1.exe"
        for ([int] $count = 2; Test-Path $oldfilename; $count++) {
            [string] $oldfilename = "vs_enterprise_$($count).exe"
        }

        Log "Renaming file: '$filename' -> '$oldfilename'"
        ren $filename $oldfilename
    }

    [string] $url = "https://download.visualstudio.microsoft.com/download/pr/f6473c9f-a5f6-4249-af28-c2fd14b6a0fb/098edaaff3de27abde17078ebf37c1f9ee64756d0e9d6d15d49a929651bb2446/vs_Enterprise.exe"

    Invoke-WebRequest $url -OutFile $filename
}

function Install-VisualStudio() {
    [string] $binary = Join-Path (pwd).Path $filename

    .\closewindow Install Error

    Log "Starting: '$binary'"
    if (Test-Path "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise") {
        Start-Process -Wait $binary "update --all --quiet --wait"
    }
    else {
        Start-Process -Wait $binary "--all --quiet --wait"
    }

    .\closewindow CloseWindow

    # Remove CVEs
    rd -Recurse -Force "C:\Python27" -ErrorAction SilentlyContinue
    rd -Recurse -Force "C:\Python27amd64" -ErrorAction SilentlyContinue
    rd -Recurse -Force "C:\Microsoft\AndroidNDK\android-ndk-r16b\prebuilt\windows\bin" -ErrorAction SilentlyContinue
    rd -Recurse -Force "C:\Microsoft\AndroidNDK64\android-ndk-r16b\prebuilt\windows-x86_64\bin" -ErrorAction SilentlyContinue
}

UpdateVS
