Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

# Re-encrypt EncryptedClientSecret using the following command:
# Read-Host | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString

Start-Transcript -Path "UpdateCert_transcript.log" -Append

function Main() {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    MakeSure-TempFolderExists

    [string] $folder = Download-Wacs

    if (Test-Path "PreAction.ps1") {
        .\PreAction.ps1
    }

    $firewallInfo = Get-Content "firewall.json" | ConvertFrom-Json
    Open-Firewall $firewallInfo

    try {
        cd $folder
        Log "Updating cert..."
        .\wacs --renew
    }
    finally {
        Close-Firewall $firewallInfo
    }

    if (Test-Path "PostAction.ps1") {
        .\PostAction.ps1
    }
}

function MakeSure-TempFolderExists() {
    [string] $tempfolder = [IO.Path]::GetTempPath()
    if (!(Test-Path $tempfolder)) {
        Log "Creating temp folder: '$tempfolder'"
        md $tempfolder | Out-Null
    }
}

function Open-Firewall($json) {
    [string] $tenantId = $json.TenantId
    [string] $subscriptionId = $json.SubscriptionId
    [string] $clientId = $json.ClientId
    $ss = $json.EncryptedClientSecret | ConvertTo-SecureString
    [string] $clientsecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss))
    [string] $resourceGroup = $json.ResourceGroup
    [string] $nsgName = $json.NsgName
    [string] $ruleName = $json.RuleName
    [string] $priority = $json.Priority

    Log "Logging in: '$tenantId', '$clientId'"
    az login --service-principal -t $tenantId -u $clientId -p $clientSecret | Out-Null

    Log "Setting subscription: '$subscriptionId'"
    az account set --subscription $subscriptionId | Out-Null

    Create-Rule $resourceGroup $nsgName $ruleName $priority

    Log "Waiting..."
    Start-Sleep 60
}

function Create-Rule([string] $resourceGroup, [string] $nsgName, [string] $ruleName, [string] $priority) {
    try {
        Log "Adding firewall rule: '$resourceGroup', '$nsgName', '$ruleName', '$priority'"
        az network nsg rule create -g $resourceGroup --nsg-name $nsgName -n $ruleName --priority $priority --access Allow --destination-address-prefixes "*" --destination-port-ranges 80 --direction Inbound --protocol Tcp --source-address-prefixes "*" --source-port-ranges "*" | Out-Null
    }
    catch {
        Log "Updating firewall rule: '$resourceGroup', '$nsgName', '$ruleName', '$priority'"
        az network nsg rule update -g $resourceGroup --nsg-name $nsgName -n $ruleName --priority $priority --access Allow --destination-address-prefixes "*" --destination-port-ranges 80 --direction Inbound --protocol Tcp --source-address-prefixes "*" --source-port-ranges "*" --set | Out-Null
    }
}

function Close-Firewall($json) {
    [string] $resourceGroup = $json.ResourceGroup
    [string] $nsgName = $json.NsgName
    [string] $ruleName = $json.RuleName

    Log "Deleting firewall rule: '$resourceGroup', '$nsgName', '$ruleName'"
    az network nsg rule delete -g $resourceGroup --nsg-name $nsgName -n $ruleName | Out-Null
}

function Download-Wacs() {
    [string] $repoUrl = "https://api.github.com/repos/PKISharp/win-acme/releases/latest"
    [string] $filePattern = "win-acme.v*.x64.pluggable.zip"
    [string] $url = ((Invoke-WebRequest $repoUrl).Content | ConvertFrom-Json).assets | ? { $_.Name -like $filePattern } | % { $_.browser_download_url }
    [int] $lastSlash = $url.LastIndexOf("/")
    [string] $filename = $url.Substring($lastSlash + 1)

    if (!(Test-Path $filename)) {
        Log "Downloading: '$url' -> '$filename'"
        Invoke-WebRequest -Uri $url -OutFile $filename
    }

    [string] $folder = Extract-Wacs $filename

    return $folder
}

function Extract-Wacs([string] $filename) {
    [string] $folder = "win-acme"
    if (Test-Path $folder) {
        Log "Deleting folder: '$folder'"
        rd -Recurse -Force $folder
        Start-Sleep 2
    }

    Log "Creating folder: '$folder'"
    md $folder | Out-Null

    Log "Extracting: '$filename' -> '$folder'"
    7z x "-o$folder" $filename | Out-Null

    return $folder
}

function Log([string] $message, $color) {
    ([DateTime]::UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + ": " + $message) | Add-Content "UpdateCert.log"

    if ($color) {
        Write-Host $message -f $color
    }
    else {
        Write-Host $message -f Green
    }
}

Main
