Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

function UpdateGit() {
    Log "Updating Git..."

    $watch = [Diagnostics.Stopwatch]::StartNew()

    [string] $filename = Download-Git
    if (!$filename) {
        return
    }

    Install-Git

    Delete-OldFiles 3

    Log "Done: $($watch.Elapsed)"
}

function Download-Git() {
    [string] $githubUrl = "https://api.github.com/repos/git-for-windows/git/releases/latest"
    [string[]] $urls = @(Invoke-WebRequest $githubUrl | ConvertFrom-Json).assets.browser_download_url | ? { $_.EndsWith("-64-bit.exe") }

    if ($urls.Count -eq 0) {
        Log "Couldn't find any file to download at: '$githubUrl'" Yellow
        return
    }
    if ($urls.Count -gt 1) {
        Log "Too many files matched at: '$githubUrl': '$($urls -join "', '")'" Yellow
        return
    }

    [string] $url = $urls[0]

    if (!$url.StartsWith("https://github.com/git-for-windows/git/releases/download/")) {
        Log "Invalid url: '$url'" Yellow
        return
    }

    [string] $filename = Split-Path -Leaf $url
    if (Test-Path $filename) {
        if ((dir $filename).Length -eq 0) {
            del $filename
        }
        else {
            Log "Git already updated with: '$filename'"
            return;
        }
    }

    Log "Downloading: '$url' -> '$filename'"
    Invoke-WebRequest $url -OutFile $filename

    return $filename
}

function Install-Git() {
    [string] $binary = Join-Path (pwd).Path $filename
    [string] $cmdargs = '/VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS /COMPONENTS="icons,ext\reg\shellhere,assoc,assoc_sh"'

    Log "Running: '$binary' '$cmdargs'"

    [Diagnostics.Process] $process = [Diagnostics.Process]::Start($binary, $cmdargs)
    $process.WaitForExit()
}

function Delete-OldFiles([int] $keep) {
    $files = @(dir "Git-*-64-bit.exe" | Sort-Object "LastWriteTime" -Descending | Select-Object -Skip $keep)
    Log "Found $($files.Count) old git install files."

    foreach ($file in $files) {
        [string] $filename = $file.FullName
        Log "Deleting: '$filename'"
        del $filename
    }
}

UpdateGit
