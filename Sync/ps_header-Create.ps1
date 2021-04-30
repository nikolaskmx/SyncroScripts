<#
.SYNOPSIS
    Thuis script will create ps_header.ps1 including all secrets vars within the same directory
.DESCRIPTION
    It connects to Microsoft and sets up custome app with correct permisions for comunicating with Azure and PartnerCenter
    It only need to be run once, for refresh token, please use Config-RefreshHeader.ps1 at least once every 90 days
    Major portion of the code is based on work of Kevin Tegelaar and script https://github.com/KelvinTegelaar/SecureAppModel/blob/master/Create-SecureAppModel.ps1
    You need to run in Escalated mode outside of ISE
.EXAMPLE
    .\Config-CreateHeader.ps1
    .\Config-CreateHeader.ps1 -DisplayName "M365 to Syncro Sync" 
    .\Config-CreateHeader.ps1 -DisplayName "M365 to Syncro Sync" -SyncroSubdomain "test.shield"
    .\Config-CreateHeader.ps1 -DisplayName "M365 to Syncro Sync" -SyncroSubdomain "test.shield" -TenantID "test.onmicrosoft.com"
.NOTES
    Author: Mariusz Sztanga
    Date:   April, 29, 2021   
#>
Param
( 
    [Parameter(Mandatory = $false)]
    [switch]$ConfigurePreconsent,
    [Parameter(Mandatory = $true)]
    [string]$DisplayName,
    [Parameter(Mandatory = $true)]
    [string]$SyncroSubdomain,
    [Parameter(Mandatory = $false)]
    [string]$TenantId
)
 
$ErrorActionPreference = "Stop"
 
###### Load modules #
$modules = @('MSOnline','AzureAD','PartnerCenter')
try {
    foreach($module in $modules) {
        Write-Host "Verifying $module Module" -ForegroundColor Green
        if (Get-Module -ListAvailable -Name $module) {
        } else {
            Install-Module -Name $module 
        }
    }
}
catch {
    # check if all modules loaded successful
    write-host "Issue with loading modules - please re-run the script" -foregroundcolor Red
    exit
}

 
try {
    Write-Host -ForegroundColor Green "When prompted please enter the appropriate credentials... Warning: Window might have pop-under in VSCode"
 
    if([string]::IsNullOrEmpty($TenantId)) {
        Connect-AzureAD | Out-Null
 
        $TenantId = $(Get-AzureADTenantDetail).ObjectId
    } else {
        Connect-AzureAD -TenantId $TenantId | Out-Null
    }
} catch [Microsoft.Azure.Common.Authentication.AadAuthenticationCanceledException] {
    # The authentication attempt was canceled by the end-user. Execution of the script should be halted.
    Write-Host -ForegroundColor Yellow "The authentication attempt was canceled. Execution of the script will be halted..."
    Exit
} catch {
    # An unexpected error has occurred. The end-user should be notified so that the appropriate action can be taken.
    Write-Error "An unexpected error has occurred. Please review the following error message and try again." `
        "$($Error[0].Exception)"
}
 
$adAppAccess = [Microsoft.Open.AzureAD.Model.RequiredResourceAccess]@{
    ResourceAppId = "00000002-0000-0000-c000-000000000000";
    ResourceAccess =
    [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
        Id = "5778995a-e1bf-45b8-affa-663a9f3f4d04";
        Type = "Role"},
    [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
        Id = "a42657d6-7f20-40e3-b6f0-cee03008a62a";
        Type = "Scope"},
    [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
        Id = "311a71cc-e848-46a1-bdf8-97ff7156d8e6";
        Type = "Scope"}
}
 
$graphAppAccess = [Microsoft.Open.AzureAD.Model.RequiredResourceAccess]@{
    ResourceAppId = "00000003-0000-0000-c000-000000000000";
    ResourceAccess =
        [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
            Id = "bf394140-e372-4bf9-a898-299cfc7564e5";
            Type = "Role"},
        [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
            Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61";
            Type = "Role"}
}
 
$partnerCenterAppAccess = [Microsoft.Open.AzureAD.Model.RequiredResourceAccess]@{
    ResourceAppId = "fa3d9a0c-3fb0-42cc-9193-47c7ecd2edbd";
    ResourceAccess =
        [Microsoft.Open.AzureAD.Model.ResourceAccess]@{
            Id = "1cebfa2a-fb4d-419e-b5f9-839b4383e05a";
            Type = "Scope"}
}
 
$SessionInfo = Get-AzureADCurrentSessionInfo

###### setting up custom AD App
Write-Host -ForegroundColor Green "Creating the Azure AD application and related resources..."
$app = New-AzureADApplication -AvailableToOtherTenants $true -DisplayName $DisplayName -IdentifierUris "https://$($SessionInfo.TenantDomain)/$((New-Guid).ToString())" -RequiredResourceAccess $adAppAccess, $graphAppAccess, $partnerCenterAppAccess -ReplyUrls @("urn:ietf:wg:oauth:2.0:oob","https://login.microsoftonline.com/organizations/oauth2/nativeclient","https://localhost","http://localhost","http://localhost:8400")
$password = New-AzureADApplicationPasswordCredential -ObjectId $app.ObjectId
$spn = New-AzureADServicePrincipal -AppId $app.AppId -DisplayName $DisplayName
$adminAgentsGroup = Get-AzureADGroup -Filter "DisplayName eq 'AdminAgents'"
Add-AzureADGroupMember -ObjectId $adminAgentsGroup.ObjectId -RefObjectId $spn.ObjectId
write-host "Sleeping for 30 seconds to allow app creation on O365" -foregroundcolor green
start-sleep 30
write-host "Please approve General consent form." -ForegroundColor Green
$PasswordToSecureString = $password.value | ConvertTo-SecureString -asPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($($app.AppId),$PasswordToSecureString)
$token = New-PartnerAccessToken -ApplicationId "$($app.AppId)" -Scopes 'https://api.partnercenter.microsoft.com/user_impersonation' -ServicePrincipal -Credential $credential -Tenant $($spn.AppOwnerTenantID) -UseAuthorizationCode

####### Setting up approval for Exchange Online
write-host "Please approve Exchange consent form." -ForegroundColor Green
Start-Process https://microsoft.com/devicelogin
$Exchangetoken = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716' -Scopes 'https://outlook.office365.com/.default' -Tenant $($spn.AppOwnerTenantID) -UseDeviceAuthentication

####### Setting up approval for PartnerCenter
write-host "PartnerCenter permissions required at https://login.microsoftonline.com/$($spn.AppOwnerTenantID)/adminConsent?client_id=$($app.AppId)"
Start-Process https://login.microsoftonline.com/$($spn.AppOwnerTenantID)/adminConsent?client_id=$($app.AppId)

write-host "Press any key after auth. An error report about incorrect URIs is expected!"
[void][System.Console]::ReadKey($true)


####### Syncro Part #######
Write-Host ""
Write-Host "Initiating request for Syncro API key"  -foregroundcolor green
Write-Host "You will be directed to new token creation page, please adjust permissions based on needs of the scripts"

Start-Process https://$SyncroSubdomain.syncromsp.com/api_tokens/new#new-api-token
$syncroAPI = Read-Host -Prompt "Enter your Syncro API key "

# Create content for file
$fileValue = @"
######### Secrets Created on $(Get-Date) #########
`$ApplicationId         = '$($app.AppId)'
`$ApplicationSecret     = '$($password.Value)' | ConvertTo-SecureString -Force -AsPlainText
`$TenantID              = '$($spn.AppOwnerTenantID)'
`$refreshToken          = '$($token.refreshtoken)'
`$ExchangeRefreshToken  = '$($ExchangeToken.Refreshtoken)'

`$SyncroAPIKey          = 'Bearer $syncroAPI'
`$SyncroSubdomain       = '$SyncroSubdomain'
######### Secrets #########

"@

$psHeaderPath = "$PSScriptRoot\ps_header.ps1"

Add-Content -Path $psHeaderPath -Value $fileValue

Write-Host ""
Write-Host "Information saved to $psHeaderPath"