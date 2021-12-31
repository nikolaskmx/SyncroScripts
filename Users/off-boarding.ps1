#########User info######
$userToOffboard = Read-Host -Prompt "Enter username (olduser@client.com)"
$CustomerDefaultDomainname = Read-Host -Prompt "Enter default domain for the client"

######### Secrets #########
$ApplicationId = 'ApplicationID'
$ApplicationSecret = 'ApplicationSecret' | ConvertTo-SecureString -Force -AsPlainText
$RefreshToken = 'RefreshToken'
$ExchangeRefreshToken = 'ExchangeToken'
$UPN = "UPN-Used-to-Generate-Tokens"

#include header with secrets or runs script to create it
if (Test-Path "$PSScriptRoot\ps_header.ps1") { 
    . "$PSScriptRoot\ps_header.ps1"
    } else {
        Write-Host  "Missing ps_header.ps1 with secret hashes" -ForegroundColor Red
        $runSetup = Read-Host -Prompt "Do you want to run ps_header-Create.ps1 (Y/N)?"
        if ($runSetup -eq 'Y') {
            $CommandLine = '-File "'+ $PSScriptRoot+ '\ps_header-Create.ps1"'
            Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine -Wait
        }
        exit
    }
######### Secrets #########
write-host "Logging in to M365 using the secure application model" -ForegroundColor Green
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $CustomerDefaultDomainname
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $CustomerDefaultDomainname
write-host "Logging into Azure AD." -ForegroundColor Green
Connect-AzureAD -AadAccessToken $aadGraphToken.AccessToken -AccountId $UPN -MsAccessToken $graphToken.AccessToken -TenantId $CustomerDefaultDomainname
write-host "Connecting to Exchange Online" -ForegroundColor Green
$token = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716'-RefreshToken $ExchangeRefreshToken -Scopes 'https://outlook.office365.com/.default' -Tenant $CustomerDefaultDomainname
$tokenValue = ConvertTo-SecureString "Bearer $($token.AccessToken)" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($upn, $tokenValue)
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell-liveid?DelegatedOrg=$($CustomerDefaultDomainname)&BasicAuthToOAuthConversion=true" -Credential $credential -Authentication Basic -AllowRedirection
Import-PSSession $session -AllowClobber
write-host "Removing users from Azure AD Groups" -ForegroundColor Green
$MemberID = (Get-AzureADUser -ObjectId $userToOffboard).objectId
Get-AzureADUserMembership -ObjectId $MemberID -All $true | Where-Object { $_.ObjectType -eq "Group" -and $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $false } | ForEach-Object { 
    write-host "    Removing using from $($_.displayname)" -ForegroundColor green
    Remove-AzureADGroupMember -ObjectId $_.ObjectID -MemberId $MemberID
}
write-host "Removing users from Unified Groups and Teams" -ForegroundColor Green
$OffboardingDN = (get-mailbox -Identity $userToOffboard -IncludeInactiveMailbox).DistinguishedName
Get-Recipient -Filter "Members -eq '$OffboardingDN'" -RecipientTypeDetails 'GroupMailbox' | foreach-object { 
    write-host "    Removing using from $($_.name)" -ForegroundColor green
    Remove-UnifiedGroupLinks -Identity $_.ExternalDirectoryObjectId -Links $userToOffboard -LinkType Member -Confirm:$false }
 
write-host "Removing users from Distribution Groups" -ForegroundColor Green
Get-Recipient -Filter "Members -eq '$OffboardingDN'" | foreach-object { 
    write-host "    Removing using from $($_.name)" -ForegroundColor green
    Remove-DistributionGroupMember -Identity $_.ExternalDirectoryObjectId -Member $OffboardingDN -BypassSecurityGroupManagerCheck -Confirm:$false }
 
write-host "Setting mailbox to Shared Mailbox" -ForegroundColor Green
Set-Mailbox $userToOffboard -Type Shared
write-host "Hiding user from GAL" -ForegroundColor Green
Set-Mailbox $userToOffboard -HiddenFromAddressListsEnabled $true
 
write-host "Removing License from user." -ForegroundColor Green
$AssignedLicensesTable = Get-AzureADUser -ObjectId $userToOffboard | Get-AzureADUserLicenseDetail | Select-Object @{n = "License"; e = { $_.SkuPartNumber } }, skuid 
if ($AssignedLicensesTable) {
    $body = @{
        addLicenses    = @()
        removeLicenses = @($AssignedLicensesTable.skuid)
    }
    Set-AzureADUserLicense -ObjectId $userToOffboard -AssignedLicenses $body
}
 
write-host "Removed licenses:"
$AssignedLicensesTable
Remove-PSSession $session