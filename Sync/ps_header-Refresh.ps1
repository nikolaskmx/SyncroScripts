<#
.SYNOPSIS
    Thuis script will create ps_header.ps1 including all secrets vars within the same directory
.DESCRIPTION
    It connects to Microsoft and refreshes tokens. Tokens will auto expire if not refreshed within 90 days
    This script should be setup to schedule run every 75 days 
    Major portion of the code is based on work of Kevin Tegelaar and script https://www.cyberdrain.com/automating-with-powershell-secure-app-model-refresh-tokens/
.EXAMPLE
    .\ps_header-Refresh.ps1
.NOTES
    Author: Mariusz Sztanga
    Date:   April, 29, 2021   
#>

#include header with secrets
if (Test-Path "$PSScriptRoot\ps_header.ps1") { 
    . "$PSScriptRoot\ps_header.ps1"
} else {
    Write-Host  "Missing ps_header.ps1 with secret hashes" -ForegroundColor Red
    $runSetup = Read-Host -Prompt "Do you want to run ps_header-Create.ps1 (Y/N)?"
    if ($runSetup -eq 'Y') {
        $CommandLine = '-File "'+ $PSScriptRoot+ '\ps_header-Create.ps1"'
        Write-Host $CommandLine
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine -Wait
    }
    exit
}

$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)
 
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID
#$graphToken    = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID
$Exchangetoken = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716' -RefreshToken $ExchangeRefreshToken -Scopes 'https://outlook.office365.com/.default' -ServicePrincipal -Tenant $tenantID

# Create content for file
$fileValue = @"
######### Secrets Refreshed #########
`$refreshDate       = '$(Get-Date)'
`$refreshToken       = '$($aadGraphToken.RefreshToken)'
`$ExchangeRefreshToken  = '$($ExchangeToken.Refreshtoken)'
######### Secrets #########

"@

$psHeaderPath = "$PSScriptRoot\ps_header.ps1"

Add-Content -Path $psHeaderPath -Value $fileValue

Write-Host ""
Write-Host "Information saved to $psHeaderPath"