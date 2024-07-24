param (
    [string]$email,
    [string]$password,
    [string]$client
)

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    if ($null -ne $email -and $null -ne $password -and $null -ne $client) {
        $arguments = "-File `"$($myinvocation.MyCommand.Definition)`" -email `"$email`" -password `"$password`" -client `"$client`""
        Start-Process powershell.exe -ArgumentList $arguments -Verb RunAs
    } else {
        $arguments = "-File `"$($myinvocation.MyCommand.Definition)`""
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
        Write-Host "After the module is installed it may exit.  Just run it again to continue..."
        Write-Host "The $ModuleName module is not installed. Installing..." -ForegroundColor Yellow
        Install-Module -Name $ModuleName
        Write-Output "Finished installing. Please restart!"
        
        exit
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
        Connect-ExchangeOnline -UserPrincipalName $username -ShowProgress $true
    }
    catch {
        Write-Host "Username/password login failed.  We will try again with a user prompted login" -ForegroundColor Yellow
        Connect-AzureAD
    }
    $users = Get-AzureADUser -All $true
    $Data = @()
    foreach ($user in $users) {
        $mfaDevices = Get-AzureADUserRegisteredDevice -ObjectId $user.ObjectId
        $hasMFA = -not ($null -eq $mfaDevices)
        $licenses = Get-AzureADSubscribedSku | Select SkuPartNumber, @{Name="ActiveUnits";Expression={$_.SkuCapacity}}, @{Name="ConsumedUnits";Expression={$_.ConsumedUnits}}
        $licenseList = $licenses.SkuPartNumber -join '; '
        $itsStatus = @{ $true = "Enabled"; $false = "Disabled"; "Unknown" = "Unknown" }
        $popEnabled = "Unknown"
        $owaEnabled = "Unknown"
        $imapEnabled = "Unknown"
        $otherEmailAppsEnabled = "Unknown"
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
    Disconnect-AzureAD

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
Test-ModuleInstallation -ModuleName "ImportExcel"
Test-ModuleInstallation -ModuleName "AzureAD.Standard.Preview"
Test-ModuleInstallation -ModuleName "ExchangeOnlineManagement"
$baseFolderPath = "\\192.168.203.207\Shared Folders\PBIData"
$excelFilePath = "$baseFolderPath\NetDoc\Manual\Admin Emails.xlsx"
$folderPath = Join-Path -Path $baseFolderPath -ChildPath "O365_data"
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
        $excelData = Import-Excel -Path $excelFilePath
    } Catch {
        Write-Host "Failed to find the admin list here: $excelFilePath"
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