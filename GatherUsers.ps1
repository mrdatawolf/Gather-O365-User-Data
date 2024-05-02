if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell.exe "-File", ($myinvocation.MyCommand.Definition) -Verb RunAs
    exit
}
function Test-ModuleInstallation {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
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
        $Data += New-Object PSObject -Property @{
            GatherDate = Get-Date -Format "MM/dd/yyyy"
            CompanyName = (Get-AzureADTenantDetail).DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Licenses = $licenseList
            HasMFA = $hasMFA
        }
    }
    Disconnect-AzureAD

    return $Data
}

Test-ModuleInstallation -ModuleName "ImportExcel"
Test-ModuleInstallation -ModuleName "AzureAD.Standard.Preview"
$baseFolderPath = "\\192.168.203.207\Shared Folders\PBIData"
$excelFilePath = "$baseFolderPath\NetDoc\Manual\Admin Emails.xlsx"
try {
    $excelData = Import-Excel -Path $excelFilePath
} Catch {
    Write-Host "Failed to find the admin list here: $excelFilePath"
    Pause
}
$Data = @()
foreach ($row in $excelData) {
    if ($row.automate -eq 1) {
        $email = $row.Email
        $password = ConvertTo-SecureString -String $row.Password -AsPlainText -Force
        Write-Host "Current admin email: $email"
        $usersData  = Get-UserData -BaseFolderPath $baseFolderPath -username $email -securePassword $password
        if($null -ne $usersData) {
            $Data += $usersData
        }
    }
} 
$dateForFileName = Get-Date -Format "MM_dd_yyyy"
$folderPath = Join-Path -Path $baseFolderPath -ChildPath "O365_data"
if (!(Test-Path -Path $folderPath)) { # Check if the folder exists, if not, create it
    New-Item -ItemType Directory -Path $folderPath | Out-Null
}
$filePath = Join-Path -Path $folderPath -ChildPath "users_$dateForFileName.csv"
Write-Host $Data
Pause
$Data | Select-Object GatherDate, CompanyName, UserPrincipalName, Licenses, HasMFA | Export-Csv -Path $filePath -NoTypeInformation

Pause

