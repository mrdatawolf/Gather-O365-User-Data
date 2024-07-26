param (
    [string]$email,
    [string]$password,
    [string]$client
)
if (-not $scriptRoot) {
    $scriptRoot = $PSScriptRoot
}
$envFilePath = Join-Path -Path $scriptRoot -ChildPath ".env"
if (Test-Path $envFilePath) {
    Get-Content $envFilePath | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]+?)\s*=\s*(.*?)\s*$") {
            [System.Environment]::SetEnvironmentVariable($matches[1], $matches[2])
        }
    }
} else {
    Write-Host "You need a .env setup before running this!" -ForegroundColor Red
    Pause
    exit
}
$baseFolderPath = [System.Environment]::GetEnvironmentVariable("BaseFolderPath")
$aEmailsFilePath = [System.Environment]::GetEnvironmentVariable("AEmailsFilePath")
$folderPath = Join-Path -Path $baseFolderPath -ChildPath "O365_data"
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    if ($null -ne $email -and $null -ne $password) {
        $arguments = "-File `"$($myinvocation.MyCommand.Definition)`" -email `"$email`" -password `"$password`" -client `"$client`" -scriptRoot `"$scriptRoot`""
        Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    } else {
        $arguments = "-File `"$($myinvocation.MyCommand.Definition)`" -scriptRoot `"$scriptRoot`""
        Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    }
    exit
}
function Test-ModuleInstallation {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "The $ModuleName module is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName -Force
        
        return $false
    } else {
        Write-Host "Importing $ModuleName..." -ForegroundColor Green
        Import-Module $ModuleName
    }

    return $true
}


function Get-UserData {
    param (
        [Parameter(Mandatory=$true)]
        [string]$BaseFolderPath,
        [string]$username,
        [SecureString]$securePassword
    )
    try {
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securePassword
        Connect-AzureAD -Credential $credential 2>$null
        Connect-ExchangeOnline -Credential $credential -ShowBanner:$false
    }
    catch {
        Write-Host "Username/password login failed.  We will try again with a user prompted login" -ForegroundColor Yellow
        Connect-AzureAD
        Connect-ExchangeOnline
    }
    $users = Get-AzureADUser -All $true
    $Data = @()
    foreach ($user in $users) {
        if ($null -ne $user) {
            # Check if the user has a mailbox
            $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
            if ($null -eq $mailbox) {
                Write-Host "No mailbox found for user $($user.UserPrincipalName)"
                continue
            }
            $mfaDevices = Get-AzureADUserRegisteredDevice -ObjectId $user.ObjectId
            $hasMFA = -not ($null -eq $mfaDevices)
            $licenses = Get-AzureADSubscribedSku | Select-Object SkuPartNumber, @{Name="ActiveUnits";Expression={$_.SkuCapacity}}, @{Name="ConsumedUnits";Expression={$_.ConsumedUnits}}
            $licenseList = $licenses.SkuPartNumber -join '; '
            $itsStatus = @{ $true = "Enabled"; $false = "Disabled"; "Unknown" = "Unknown" }
            $casMailbox = Get-CASMailbox -Identity $user.UserPrincipalName
            $popEnabled = $casMailbox.POPEnabled
            $owaEnabled = $casMailbox.OWAEnabled
            $imapEnabled = $casMailbox.IMAPEnabled
            $otherEmailAppsEnabled = $casMailbox.ActiveSyncEnabled
            $Data += New-Object PSObject -Property @{
                GatherDate = Get-Date -Format "MM/dd/yyyy"
                CompanyName = (Get-AzureADTenantDetail).DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Licenses = $licenseList
                MFA = $itsStatus[$hasMFA]
                POP = $itsStatus[$popEnabled]
                OWA = $itsStatus[$owaEnabled]
                IMAP = $itsStatus[$imapEnabled]
                OtherEmailApps = $itsStatus[$otherEmailAppsEnabled]
            }
        }
    }
    
    Disconnect-AzureAD
    Disconnect-ExchangeOnline -Confirm:$false

    return $Data
}

function Export-UserData {
    param (
        [string]$email,
        [SecureString]$password,
        [string]$clientName,
        [string]$baseFolderPath,
        [string]$dateForFileName
    )

    $Data = @()
    $usersData  = Get-UserData -BaseFolderPath $baseFolderPath -username $email -securePassword $password
    if($null -ne $usersData) {
        $Data += $usersData
    }
    $filePath = Join-Path -Path $folderPath -ChildPath ("{0}_users_{1}.csv" -f $clientName, $dateForFileName)
    $Data | Select-Object GatherDate, CompanyName, UserPrincipalName, Licenses, MFA, POP, OWA, IMAP, OtherEmailApps | Export-Csv -Path $filePath -NoTypeInformation
    Write-Host "Finished writting the file for $clientName."
}

$dateForFileName = Get-Date -Format "MM_dd_yyyy"
$modules = @("ImportExcel", "AzureAD.Standard.Preview", "ExchangeOnlineManagement")
foreach ($module in $modules) {
    $result = Test-ModuleInstallation -ModuleName $module
    if (-not $result) {
        Write-Host "Please restart the script after installing the required modules." -ForegroundColor Red
        exit
    }
}
Write-Host "All required modules are installed and imported."
if (!(Test-Path -Path $folderPath)) { # Check if the folder exists, if not, create it
    New-Item -ItemType Directory -Path $folderPath | Out-Null
}
if ($email -and $password) {
    $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    Write-Host "Current admin email: $email"
    $clientName = if($client) { $client } else { "test" }
    Export-UserData -email $email -password $securePassword -clientName $clientName -baseFolderPath $baseFolderPath -dateForFileName $dateForFileName
} else {
    try {
        $excelData = Import-Excel -Path $aEmailsFilePath
    } Catch {
        Write-Host "Failed to find the admin list here: $aEmailsFilePath"
        Pause
    }
    foreach ($row in $excelData) {
        if ($row.automate -eq 1) {
            $email = $row.Email
            $securePassword = ConvertTo-SecureString -String $row.Password -AsPlainText -Force
            Write-Host "Current admin email: $email"
            $clientName = $row.Client
            Export-UserData -email $email -password $securePassword -clientName $clientName -baseFolderPath $baseFolderPath -dateForFileName $dateForFileName
        }
    } 
}
Write-Host "Completed."
Pause