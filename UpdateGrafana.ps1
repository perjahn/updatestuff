Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

function Main([string[]] $mainargs) {
    [Diagnostics.Stopwatch] $totalwatch = [Diagnostics.Stopwatch]::StartNew()

    [int] $zipfileSize = 50mb

    [string] $downloadUrl = Get-DownloadUrl

    [string] $zipfile = Split-Path -Leaf $downloadUrl
    if ((Test-Path $zipfile) -and (dir $zipfile).Length -ge $zipfileSize) {
        Log "Grafana already updated with: '$zipfile'"
        exit 0
    }

    Robust-Download $downloadUrl $zipfile $zipfileSize

    Install-Grafana $zipfile

    Delete-OldFiles "zip" 3

    Log "Done: $($totalwatch.Elapsed)" Cyan
}

function Get-DownloadUrl() {
    [string] $pageurl = "https://grafana.com/grafana/download?platform=windows"

    #<a href="https://dl.grafana.com/oss/release/grafana-x.y.z.windows-amd64.zip">Download the zip file</a>

    Log "Downloading url: '$pageurl'"

    [string] $page = Invoke-WebRequest $pageurl
    [string[]] $rows = @($page.Split("`n"))
    Log "Got $($rows.Count) rows."

    for ([int] $i = 0; $i -lt $rows.Count; $i++) {
        [string] $row = $rows[$i]
        [string] $filename = "windows-amd64.zip"
        [int] $end = $row.IndexOf($filename)
        if ($end -ne -1) {
            [int] $start = $row.LastIndexOf("https://", $end)
            if ($end -ne -1) {
                [string] $url = $row.Substring($start, $end - $start + $filename.Length)
                Log "Got url: '$url'"

                return $url
            }
        }
    }
}

function Robust-Download([string] $url, [string] $outfile, [int] $fileSize) {
    if (!$url.StartsWith("https://dl.grafana.com/")) {
        LogError "Invalid url: '$url'"
        exit 1
    }

    for ([int] $tries = 1; !(Test-Path $outfile) -or (dir $outfile).Length -lt $fileSize; $tries++) {
        if (Test-Path $outfile) {
            Log "Deleting (try $tries): '$outfile'"
            del $outfile
        }

        Log "Downloading (try $tries): '$url' -> '$outfile'"
        try {
            Invoke-WebRequest $url -OutFile $outfile
        }
        catch {
            Log "Couldn't download (try $tries): '$url' -> '$outfile'"
            Start-Sleep 5
        }

        if (!(Test-Path $outfile) -or (dir $outfile).Length -lt $fileSize) {
            if ($tries -lt 10) {
                Log "Couldn't download (try $tries): '$url' -> '$outfile'"
            }
            else {
                Log "Couldn't download (try $tries): '$url' -> '$outfile'"
                exit 1
            }
        }
    }

    Log "Downloaded: '$outfile'"
}

function Install-Grafana([string] $zipfile) {
    [string] $folder = [IO.Path]::GetFileNameWithoutExtension($zipfile)

    if (Test-Path $folder) {
        [int] $retries = 1
        do {
            Log "Deleting folder (try $retries): '$folder'"
            rd -Recurse -Force $folder -ErrorAction SilentlyContinue
            if (Test-Path $folder) {
                Start-Sleep 2
                $retries++
            }
            else {
                $retries = 11
            }
        } while ($retries -le 10)
    }

    [string] $serviceName = "grafana"
    [string] $logserviceName = "filebeat"

    Log "Extracting: '$zipfile'"
    Expand-Archive $zipfile

    $subfolders = @(dir $folder -Directory)
    if ($subfolders.Count -ne 1) {
        Log "Couldn't find unique subfolder, found $($subfolders.Count) subfolders: '$($subfolders -join "', '")'"
        exit 1
    }

    [string] $subfolder = Join-Path $folder $subfolders[0].Name
    [string] $newsubfolder = "grafana"
    Log "Renaming folder: '$subfolder' -> '$newsubfolder'"
    ren $subfolder $newsubfolder
    [string] $subfolder = Join-Path $folder $newsubfolder

    [string] $installfolder = "C:\grafana"

    if (Test-Path $installfolder) {
        [string] $configfile = Join-Path $subfolder "conf" "defaults.ini"
        [string] $backupfile = Join-Path $subfolder "conf" "defaults_old.ini"
        Log "Copying file: '$configfile' -> '$backupfile'"
        copy $configfile $backupfile

        [string] $previousconfigfile = Join-Path $installfolder "conf" "defaults.ini"
        Log "Updating: '$previousconfigfile' -> '$configfile'"
        Update-Config $configfile $previousconfigfile

        [string] $exefile = Join-Path $installfolder "bin" "nssm.exe"
        [string] $binfolder = Join-Path $subfolder "bin"
        Log "Copying file: '$exefile' -> '$binfolder'"
        copy $exefile $binfolder

        [string] $datafolder = Join-Path $installfolder "data"
        Log "Copying folder: '$datafolder' -> '$subfolder'"
        copy $datafolder $subfolder

        try {
            Log "Stopping service: '$serviceName'"
            Stop-Service $serviceName
            Log "Stopping service: '$logserviceName'"
            Stop-Service $logserviceName


            [string] $source = Join-Path $installfolder "data" "*"
            [string] $targetfolder = Join-Path $subfolder "data"
            if (!(Test-Path $targetfolder)) {
                Log "Creating folder: '$targetfolder'"
                md $targetfolder | Out-Null
            }
            Log "Copying: '$source' -> '$targetfolder'"
            copy -Recurse $source $targetfolder

            [int] $retries = 1
            [string] $backupfolder = "grafana_" + (Get-Date).ToString("yyyyMMdd_HHmmss")
            do {
                Log "Renaming folder (try $retries): '$installfolder' -> '$backupfolder'"
                ren $installfolder $backupfolder -ErrorAction SilentlyContinue
                if (Test-Path $installfolder) {
                    Start-Sleep 2
                    $retries++
                }
                else {
                    $retries = 11
                }
            } while ($retries -le 10)

            [string] $rootfolder = Split-Path $installfolder
            Log "Moving folder: '$subfolder' -> '$rootfolder'"
            move $subfolder $rootfolder
        }
        finally {
            Log "Starting service: '$serviceName'"
            Start-Service $serviceName
            Log "Starting service: '$logserviceName'"
            Start-Service $logserviceName
        }
    }
    else {
        [string] $nssmurl = "https://nssm.cc/release/nssm-2.24.zip"
        [string] $zipfile = "nssm.zip"
        Invoke-WebRequest $nssmurl -OutFile $zipfile
        Expand-Archive $zipfile

        [string] $exefile = Join-Path "nssm-2.24" "win64" "nssm.exe"
        [string] $binfolder = Join-Path $subfolder "bin"
        Log "Copying file: '$exefile' -> '$binfolder'"
        copy $exefile $binfolder

        [string] $sourcefolder = $subfolder
        [string] $targetfolder = "C:\"
        Log "Moving folder: '$sourcefolder' -> '$targetfolder'"
        move $sourcefolder $targetfolder

        Log "Starting service: '$serviceName'"
        Start-Service $serviceName
    }

    Log "Deleting folder: '$folder'"
    rd $folder
}

function Update-Config([string] $filename, [string] $oldfilename) {
    Log "Reading: '$oldfilename'"
    [string[]] $rows = Get-Content $oldfilename
    Log "Got $($rows.Length) rows."

    [string] $section = ""

    [string] $password = Generate-AlphanumericPassword 24
    if ($env:grafanadomain) {
        [string] $domain = $env:grafanadomain
    }
    else {
        [string] $domain = ""
    }
    if ($env:grafanaroot_url) {
        [string] $root_url = $env:grafanaroot_url
    }
    else {
        [string] $root_url = ""
    }
    if ($env:grafanalogformat) {
        [string] $logformat = $env:grafanalogformat
    }
    else {
        [string] $logformat = ""
    }

    for ([int] $i = 0; $i -lt $rows.Length; $i++) {
        [string] $row = $rows[$i]

        if ($row.StartsWith("[") -and $row.EndsWith("]")) {
            $section = $row.Substring(1, $row.Length - 2)
            continue
        }

        [string] $settingname = "domain ="
        if ($section -eq "server" -and $row.StartsWith($settingname)) {
            $domain = $row.Substring($settingname.Length).Trim()
            Log "Found old domain: '$domain'"
        }

        [string] $settingname = "root_url ="
        if ($section -eq "server" -and $row.StartsWith($settingname)) {
            $root_url = $row.Substring($settingname.Length).Trim()
            Log "Found old root_url: '$root_url'"
        }

        [string] $settingname = "admin_password ="
        if ($section -eq "security" -and $row.StartsWith($settingname)) {
            $password = $row.Substring($settingname.Length).Trim()
            Log "Found old admin_password: $("*" * $password.Length)"
        }

        [string] $settingname = "format ="
        if ($section -eq "log.file" -and $row.StartsWith($settingname)) {
            $logformat = $row.Substring($settingname.Length).Trim()
            Log "Found old logformat: '$logformat'"
        }
    }


    Log "Reading: '$filename'"
    [string[]] $rows = Get-Content $filename
    Log "Got $($rows.Length) rows."

    [string] $section = ""

    for ([int] $i = 0; $i -lt $rows.Length; $i++) {
        [string] $row = $rows[$i]

        if ($row.StartsWith("[") -and $row.EndsWith("]")) {
            $section = $row.Substring(1, $row.Length - 2)
            continue
        }

        if ($section -eq "server" -and $row.StartsWith("domain =")) {
            $rows[$i] = "domain = " + $domain
        }
        if ($section -eq "server" -and $row.StartsWith("root_url =")) {
            $rows[$i] = "root_url = " + $root_url
        }
        if ($section -eq "remote_cache" -and $row.StartsWith("connstr =")) {
            $rows[$i] = ";connstr ="
        }
        if ($section -eq "analytics" -and $row.StartsWith("reporting_enabled =")) {
            $rows[$i] = "reporting_enabled = false"
        }
        if ($section -eq "security" -and $row.StartsWith("admin_password =")) {
            $rows[$i] = "admin_password = " + $password
        }
        if ($section -eq "log.file" -and $row.StartsWith("format =")) {
            $rows[$i] = "format = " + $logformat
        }
    }

    Log "Saving $($rows.Length) rows to: '$filename'"
    Set-Content $filename $rows
}

function Generate-AlphanumericPassword([int] $numberOfChars) {
    [char[]] $validChars = 'a'..'z' + 'A'..'Z' + [char]'0'..[char]'9'
    [string] $password = ""
    do {
        [string] $password = (1..$numberOfChars | % { $validChars[(Get-Random -Maximum $validChars.Length)] }) -join ""
    }
    while (
        !($password | ? { ($_.ToCharArray() | ? { [Char]::IsUpper($_) }) }) -or
        !($password | ? { ($_.ToCharArray() | ? { [Char]::IsLower($_) }) }) -or
        !($password | ? { ($_.ToCharArray() | ? { [Char]::IsDigit($_) }) }));

    return $password
}

function Delete-OldFiles([string] $extension, [int] $keep) {
    $files = @(dir ("*." + $extension) | Sort-Object "LastWriteTime" -Descending | Select-Object -Skip $keep)

    Log "Found $($files.Count) old $extension files."
    foreach ($file in $files) {
        [string] $filename = $file.FullName
        Log "Deleting: '$filename'"
        del $filename
    }
}

function Log([string] $message, $color) {
    [string] $annotatedMessage = [DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + ": " + $message

    $annotatedMessage | Add-Content "Grafana.log"

    if ($color) {
        Write-Host ($message) -f $color
    }
    else {
        Write-Host ($message) -f Green
    }
}

Main $args
