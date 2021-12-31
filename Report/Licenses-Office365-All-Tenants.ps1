<#
.SYNOPSIS
    This script will grab licenses from Office365 and create ticket in Syncro with usernames and usage
.DESCRIPTION
    It connects to Microsoft Azure and compares data to Syncro Contacts based on Comapany names. Contacts not in Syncro will be created
    Existing contacts will be compared to Syncro entries. If field exists in Office365 - it will be replicated to Syncro
    If Office365 field is empty and Syncro has information - this field will be preserved.
    Any contact that is not in Office365 will be removed from Syncro unless you change varible $deleteUser to $false and user will get added (D) at from of their name
    First time you run it as Admin, you will get the option to setup schedule to run it everyday at 6:35am (you can change it to your liking)

    Required permissions from Syncro:
        Contacts - Import
        Customers - Create / List/Search / View Detail / Edit / Delete
.EXAMPLE
    .\Licenses-Office365-All-Tenants.ps1
.NOTES
    Author: Mariusz Sztanga
    Date:   May, 26, 2021   
#>

# Params
$ScriptName     = $MyInvocation.MyCommand.Name
$createTicket     = $false

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

<# Setup schedule 
$taskExists = Get-ScheduledTask | Where-Object {$_.TaskName -like $ScriptName}

if($taskExists) {
    if ((Get-ScheduledTask -TaskName $ScriptName).state -like "Disabled") {Write-Host "Schedule is disabled for this script... Use Scheduler to re-enable" -ForegroundColor Red}
        
} else {
    $securePwd = Read-Host "Enter password to schedule under $($env:UserName) (leave blank to disable task):" -AsSecureString
    $taskPrincipal = New-ScheduledTaskPrincipal -UserId $env:UserName -RunLevel Highest
    $taskAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-File $($PSScriptRoot+'\'+$ScriptName+'.ps1')"
    $taskTrigger = New-ScheduledTaskTrigger -Daily -At 6:35am
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
#>


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

write-host "Connecting to MSOLService" -ForegroundColor Green
Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken
write-host "Grabbing client list" -ForegroundColor Green
$customers = Get-MsolPartnerContract -All
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
if (-not (Test-Path -LiteralPath $($ScriptDir+"\reports"))) { New-Item -Path $($ScriptDir+"\reports") -ItemType Directory | Out-Null}
$reportdate = Get-Date -format "yyyy-MM-dd"
write-host "Connecting to clients" -ForegroundColor Green

foreach ($customer in $customers) {
    $CustomerToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -Tenant $customer.TenantID
    $headers = @{ "Authorization" = "Bearer $($CustomerToken.AccessToken)" }


    write-host "Gathering Reports for $($Customer.name)" -ForegroundColor Green
    #Gathers which users currently use email and the details for these Users
    #$EmailReportsURI = "https://graph.microsoft.com/beta/reports/getEmailActivityUserDetail(period='D30')"
    #$EmailReports = (Invoke-RestMethod -Uri $EmailReportsURI -Headers $Header -Method Get -ContentType "application/json") -replace "ï»¿", "" | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Email users Report</h1>"| Out-String
    $EmailReports = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getEmailActivityUserDetail(period='D30')" -Headers $headers -Method Get -ContentType "application/json") | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Email users Report</h1>"| Out-String
    $MailboxUsage = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getMailboxUsageDetail(period='D30')" -Headers $headers -Method Get -ContentType "application/json") | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Email storage report</h1>"| Out-String
    $O365ActivationsReports = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getOffice365ActivationsUserDetail" -Headers $headers -Method Get -ContentType "application/json") | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>O365 Activation report</h1>"| Out-String
    $OneDriveActivityReports = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getOneDriveActivityUserDetail(period='D30')" -Headers $headers -Method Get -ContentType "application/json") | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>OneDrive usage report</h1>"| Out-String
    $SharepointUsageReports = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getSharePointSiteUsageDetail(period='D30')" -Headers $headers -Method Get -ContentType "application/json") | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Sharepoint usage report</h1>"| Out-String
    $TeamsDeviceReports = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getTeamsDeviceUsageUserDetail(period='D30')" -Headers $headers -Method Get -ContentType "application/json") | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Teams device report</h1>"| Out-String
    $TeamsUserReports = (Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/reports/getTeamsUserActivityUserDetail(period='D30')" -Headers $headers -Method Get -ContentType "application/json") | ConvertFrom-Csv | ConvertTo-Html -fragment -PreContent "<h1>Teams user report</h1>"| Out-String

$head = 
@"
      <Title>O365 Reports</Title>
    <style>
    body { background-color:#E5E4E2;
          font-family:Monospace;
          font-size:10pt; }
    td, th { border:0px solid black; 
            border-collapse:collapse;
            white-space:pre; }
    th { color:white;
        background-color:black; }
    table, tr, td, th {
         padding: 2px; 
         margin: 0px;
         white-space:pre; }
    tr:nth-child(odd) {background-color: lightgray}
    table { width:95%;margin-left:5px; margin-bottom:20px; }
    h2 {
    font-family:Tahoma;
    color:#6D7B8D;
    }
    .footer 
    { color:green; 
     margin-left:10px; 
     font-family:Tahoma;
     font-size:8pt;
     font-style:italic;
    }
    </style>
"@
 
$head,$EmailReports,$MailboxUsage,$O365ActivationsReports,$TeamsDeviceReports,$TeamsUserReports,$OneDriveUsageReports,$OneDriveActivityReports,$SharepointUsageReports | out-file "$ScriptDir\reports\$($reportdate) $($Customer.DefaultDomainName).html"

}