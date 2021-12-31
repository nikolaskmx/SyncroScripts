<#
.SYNOPSIS
    This script will sync contacts from Office365 to Syncro
.DESCRIPTION
    It connects to Microsoft Azure and compares data to Syncro Contacts based on Comapany names. Contacts not in Syncro will be created
    Existing contacts will be compared to Syncro entries. If field exists in Office365 - it will be replicated to Syncro
    If Office365 field is empty and Syncro has information - this field will be preserved.
    Any contact that is not in Office365 will be removed from Syncro unless you change varible $deleteUser to $false and user will get added (D) at from of their name
    First time you run it as Admin, you will get the option to setup schedule to run it everyday at 6:45am (you can change it to your liking)

    Required permissions from Syncro:
        Contacts - Import
        Customers - Create / List/Search / View Detail / Edit / Delete
.EXAMPLE
    .\Contacts-Office365-All-Tenants.ps1
.NOTES
    Author: Mariusz Sztanga
    Date:   April, 29, 2021   
#>

# Params
$ScriptName     = $MyInvocation.MyCommand.Name
$deleteUser     = $false

#debug
$debug          = $false
$logPath        = "$PSScriptRoot\$ScriptName.log" 



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

# Setup schedule 
$taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -like $ScriptName}

if($taskExists) {
    if ((Get-ScheduledTask -TaskName $ScriptName).state -like "Disabled") {Write-Host "Schedule is disabled for this script... Use Scheduler to re-enable" -ForegroundColor Red}
        
} else {
    $securePwd = Read-Host "Enter password to schedule under $($env:UserName) (leave blank to disable task):" -AsSecureString
    $taskPrincipal = New-ScheduledTaskPrincipal -UserId $env:UserName -RunLevel Highest
    $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-File $($PSScriptRoot+'\'+$ScriptName+'.ps1')"
    $taskTrigger = New-ScheduledTaskTrigger -Daily -At 6:45am
    Register-ScheduledTask -TaskName $ScriptName -Action $taskAction -Trigger $taskTrigger -Description "script to run $ScriptName" -Principal $taskPrincipal | Out-Null
    try {
        Set-ScheduledTask -TaskName $ScriptName -User $taskPrincipal.UserID -Password $([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePwd))) | Out-Null
    } catch {
        if (([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePwd))) -eq "") { 
            Disable-ScheduledTask -TaskName $ScriptName | Out-Null
        } else {
            Write-Host  "Incorrect password" -ForegroundColor Red
        }
    }
    Write-Host  "Schedule $ScriptName $((Get-ScheduledTask -TaskName $ScriptName).state)" -ForegroundColor Green
}


# check for modules and install/get if needed
$modules = @('MSOnline','AzureAD','PartnerCenter')
try {
    foreach($module in $modules) {
		Write-Host "Verifying $module Module" -ForegroundColor Green
		    if (Get-Module -ListAvailable -Name $module) {
		} 
		else {
			Install-Module -Name $module 
		}
    }
}
catch {
	Write-Host "Issue with loading Module" -ForegroundColor Red
    Exit
}

# Connect to Azure and Graph #
$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)
$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default'
 
Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken
$customers = Get-MsolPartnerContract

foreach ($customer in $customers) {
    $CustomerToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -Tenant $customer.TenantID
    $headers = @{ "Authorization" = "Bearer $($CustomerToken.AccessToken)" }
    
    $query = [System.Web.HTTPUtility]::UrlEncode($($Customer.Name))
    $companySyncroID = (Invoke-RestMethod -Uri "https://$SyncroSubdomain.syncromsp.com/api/v1/customers?business_name=$query" -Method Get -Header @{ "Authorization" = $SyncroAPIKey } -ContentType "application/json")[0].customers.id

    if ($null -eq $companySyncroID) {
        Write-host "Client $($Customer.Name) not found in SyncroMSP" -ForegroundColor Red 
    } else {
        write-host "Getting client ID# $companySyncroID for $($Customer.Name) from SyncroMSP" -ForegroundColor Green 
        # $domains = Get-MsolDomain
        $allusersO365 = (Invoke-RestMethod -Uri 'https://graph.microsoft.com/beta/users?$top=999' -Headers $Headers -Method Get -ContentType "application/json").value | Where-Object {$_.mail -ne $null -and $_.assignedLicenses -ne $null}
        $Syncro = (Invoke-RestMethod -Method Get -Uri "https://$SyncroSubdomain.syncromsp.com/api/v1/contacts?customer_id=$companySyncroID" -Header @{ "Authorization" = $SyncroAPIKey } -ContentType "application/json")
        $allusersSyncro = $syncro.contacts

        $totalPages = $Syncro.meta.total_pages
        if ($totalPages -ne 1) {
            for($i=2; $i -le $totalPages; $i++){
                $allusersSyncro += (Invoke-RestMethod -Method Get -Uri "https://$SyncroSubdomain.syncromsp.com/api/v1/contacts?customer_id=$companySyncroID&page=$i" -Header @{ "Authorization" = $SyncroAPIKey } -ContentType "application/json").contacts
            }
        }

        ### Search for users in Syncro and compare with Office365 ####
        foreach ($userO365 in $allusersO365) {
            $userSyncro = $allusersSyncro | Where-Object{$_.email -like $($userO365.mail)}
            if ($userSyncro -ne $null){
                ### Check if AD field is empty, if yes - don't overwrite it in Syncro ###
                $changed = @()
                if ($userO365.displayName -ne $null -and $userO365.displayName -ne $userSyncro.name) {$userSyncro.name = $userO365.displayName; $changed += $userO365.displayName}
                if ($userO365.streetAddress -ne $null -and $userO365.streetAddress -ne $userSyncro.address1) {$userSyncro.address1 = $userO365.streetAddress; $changed += $userO365.streetAddress}
                if ($userO365.city -ne $null -and $userO365.city -ne $userSyncro.city) {$userSyncro.city = $userO365.city; $changed += $userO365.city}
                if ($userO365.postalCode -ne $null -and $userO365.postalCode -ne $userSyncro.zip) {$userSyncro.zip = $userO365.postalCode; $changed += $userO365.postalCode}
                if ($userO365.state -ne $null -and $userO365.state -ne $userSyncro.state) {$userSyncro.state = $userO365.state; $changed += $userO365.state}
                if ($userO365.businessPhones[0] -ne $null -and $userO365.businessPhones[0] -ne $userSyncro.phone) {$userSyncro.phone = $userO365.businessPhones[0]; $changed += $userO365.businessPhones[0]}
                if ($userO365.mobilePhone -ne $null -and $userO365.mobilePhone -ne $userSyncro.mobile) {$userSyncro.mobile = $userO365.mobilePhone; $changed += $userO365.mobilePhone}
                if ($userO365.streetAddress -ne $null -and $userO365.streetAddress -ne $userSyncro.address1) {$userSyncro.address1 = $userO365.streetAddress; $changed += $userO365.streetAddress}
                if ($userO365.jobTitle -ne $null -and $userO365.jobTitle -ne $($userSyncro.properties.title)) {
                    $properties = [PSCustomObject]@{ 
                        'title'                  = $userO365.jobTitle
                    }
                    $userSyncro.properties = $properties
                    $changed += $userO365.jobTitle
                }

                if ($changed -ne $null) {   
                    Write-Host "Contact updated for $($userSyncro.name) - $changed"
                    $UpdateJSON = ConvertTo-Json $userSyncro
                    $status += $UpdateJSON
                    $status += Invoke-RestMethod -Method PUT -Uri "https://$SyncroSubdomain.syncromsp.com/api/v1/contacts/$($userSyncro.id)" -Header @{ "Authorization" = $SyncroAPIKey } -ContentType "application/json" -Body $UpdateJSON
                }
                
            } else {
                #### Found new user - adding to Syncro ####
                Write-Host "$($userO365.displayname) not found in SyncroMSP - Creating...." -ForegroundColor Red

                $properties = [PSCustomObject]@{ 
                    'title'                  = $userO365.jobTitle
                }
                $newSyncroUser = [PSCustomObject]@{
                    'customer_id' = $companySyncroID
                    'name'        = $userO365.displayname
                    'address1'    = $userO365.streetAddress
                    'city'        = $userO365.city
                    'state'       = $userO365.state
                    'zip'         = $userO365.postalCode
                    'email'       = $userO365.mail
                    'phone'       = $userO365.businessPhones[0]
                    'mobile'      = $userO365.mobilePhone
                    'properties'  = $properties
                    'opt_out'     = 'False'
                }
                
                $newUserJson = $newSyncroUser | ConvertTo-Json
                $status += Invoke-RestMethod -Method POST -Uri "https://$SyncroSubdomain.syncromsp.com/api/v1/contacts" -Header @{ "Authorization" = $SyncroAPIKey } -ContentType "application/json" -Body $newUserJson     
            }
     
        }

        #### Search for contacts in Syncro that no longer exists in Office365 ####
        foreach ($userSyncro in $allusersSyncro) {
            $userO365 = $allusersO365 | Where-Object{$_.mail -like $userSyncro.email}
            if ($null -eq $userO365){
                if ($deleteUser) {
                    Write-Host "$($userSyncro.name) was not found in Office365 - Deleting...." -ForegroundColor Red  
                    $status += Invoke-RestMethod -Method DELETE -Uri "https://$SyncroSubdomain.syncromsp.com/api/v1/contacts/$($userSyncro.id)" -Header @{ "Authorization" = $SyncroAPIKey } -ContentType "*/*" 
                } else {
                    if (-Not ($userSyncro.name -like "*[Deleted]")) {
                        $userSyncro.name = $userSyncro.name + " [Deleted]"
                        $userSyncro.notes = "User marked as deleted on $(Get-Date)"
                        $status += Invoke-RestMethod -Method PUT -Uri "https://$SyncroSubdomain.syncromsp.com/api/v1/contacts/$($userSyncro.id)" -Header @{ "Authorization" = $SyncroAPIKey } -ContentType "application/json" -Body (ConvertTo-Json $userSyncro)    
                    }
               }
            }
        }
    }
}

# log in the case if you need troubleshooting 
if ($logPath -ne $null -And $debug) {
    Write-Host "Saving..."$logPath
    Add-Content -Path $logPath -Value $(ConvertTo-Json $status)
}