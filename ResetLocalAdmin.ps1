<#
.SYNOPSIS
    Thuis script will create ps_header.ps1 including all secrets vars within the same directory
.DESCRIPTION
    It connects to Microsoft and refreshes tokens. Tokens will auto expire if not refreshed within 90 days
    This script should be setup to schedule run every 75 days 
    Major portion of the code is based on work of Kevin Tegelaar and script https://www.cyberdrain.com/automating-with-powershell-secure-app-model-refresh-tokens/
.PARAMETERS
    $numWords - number of words to add for complexity after prefix
    $username - your local admin username and name of the field in Syncro where new password will be saved
    $password - initial perfix for the password, this adds additional complexivity to password
.NOTES
    Author: Mariusz Sztanga
    Date:   April, 29, 2021   
#>

Import-Module $env:SyncroModule -WarningAction SilentlyContinue

# -----------------------------------------------------------
# Setting up local username with random password
$numWords = 3
$Username = "iavitAdmin"
$Password = "prefix@"

$Header = "key","word"
$WebResponse = $(Invoke-WebRequest "https://www.interactiveavit.com/clients/iavit/eff_short_wordlist.txt" -UseBasicParsing).Content | ConvertFrom-Csv -Delimiter "`t" -Header $Header
for ( $i = 0; $i -lt $numWords; $i++ ){
    $randWord = $WebResponse | Get-Random
    $Password = $Password + (Get-Culture).TextInfo.ToTitleCase($randWord.word)
}

$KeyPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"

$adsi = [ADSI]"WinNT://$env:COMPUTERNAME"
$group = "Administrators"
$existing = $adsi.Children | where {$_.SchemaClassName -eq 'user' -and $_.Name -eq $Username }

if ($existing -eq $null) {
    Write-Host "Creating new local user $Username."
    & NET USER $Username $Password /add /y /expires:never | Out-Null
    
    Write-Host "Adding local user $Username to $group."
    & NET LOCALGROUP $group $Username /add | Out-Null
    
    Write-Host "Ensuring password for $Username never expires."
    & WMIC USERACCOUNT WHERE "Name='$Username'" SET PasswordExpires=FALSE | Out-Null
} else {
    Write-Host "Resetting password for existing local user $Username."
    $existing.SetPassword($Password)
}

if ((Get-ItemProperty "$KeyPath\SpecialAccounts\UserList").$Username -eq $null){
    New-Item -Path "$KeyPath" -Name SpecialAccounts | Out-Null 
    New-Item -Path "$KeyPath\SpecialAccounts" -Name UserList | Out-Null
    New-ItemProperty -Path "$KeyPath\SpecialAccounts\UserList" -Name $Username -Value 0 -PropertyType DWord | Out-Null
}

#save local password back to Syncro
$setField = Set-Asset-Field -Name $Username -Value ($Password)
$logActivity = Log-Activity -Message "Password for $Username was updated" -EventName "Password Reset"