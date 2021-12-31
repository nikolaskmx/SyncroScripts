<#
.SYNOPSIS
    This script will check hardware components in ESXi via SyncroRMM
.DESCRIPTION
    Syncro script to check on ESXi hardware and report any issues via alerts in SyncroRMM
    Run this script as often as you need it via Syncro scheduler
    Create user on ESXi with Read-Only rights to access the host and enter them into Husername and Hpassword (do not use root account)
    Enter hostname(s) or IP adress(es) in $VMHosts that you want to query 
    After executing the script, you will get alerts setup in Syncro for any warnings or failures with details about component and status
    Optional CSV file with whole report can be uploaded to the computer running the script
    Initial RunTime for script should be set to at least 10 min if you don't have VMware.PowerCLI installed
.EXAMPLE
    Copy script to SyncroRMM and run it on schedule 
.NOTES
    Author: Mariusz Sztanga
    Date:   December 30, 2021   
#>

# Syncro Script Variables 
# $VMHosts        = "192.168.1.21 192.168.1.22 myVMhost" - space separated IPs or hostnames
# $Husername      = "read-only-user"
# $Hpassword      = "password-for-Husername"

$uploadLog        = $true

Import-Module $env:SyncroModule -WarningAction SilentlyContinue

try {
    $moduleCLI = "VMware.PowerCLI"
    if (-not (Get-Module -ListAvailable -Name $moduleCLI)) { 
        Install-Module -Name $moduleCLI -Force | Out-Null 
    }
} catch {
	Write-Host "Issue with loading VMware.PowerCLI Module" -ForegroundColor Red
    Exit 1
}

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -ParticipateInCEIP:$false | Out-Null

$OutArr = @()
$VMHostsArray = $VMHosts.Split(" ")
foreach($VMHost in $VMHostsArray) {
    try {
       Connect-VIServer -Server $VMHost -Protocol https -User $Husername -Password $Hpassword -ErrorAction Stop | Out-Null
    } catch [Exception]{
        $exception = $_.Exception
        Rmm-Alert -Category 'ESXi Alert - Hardware' -Body $exception.message | Out-Null
        Write-Host $($exception.message) -ForegroundColor Red
    }
    if ($null -ne $global:DefaultVIServer) {
        try {
            $VMHostObj = Get-VMHost -Name $VMHost -EA Stop | Get-View
            $sensors = $VMHostObj.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo
            foreach($sensor in $sensors){
                $object = New-Object -TypeName PSObject -Property @{
                    TimeStamp = Get-Date -format "yyyy-MMM-dd HH:mm:ss"
                    VMHost = $VMHost
                    SensorName = $sensor.Name
                    Status = $sensor.HealthState.Key
                    CurrentReading = $sensor.CurrentReading 
                } | Select-Object VMHost, SensorName, CurrentReading, Status
                $OutArr += $object
                if ($sensor.HealthState.Key -ne 'green') {
                    $alertBody = "$($sensor.Name) has issues - Reading $($sensor.CurrentReading)"
                    Rmm-Alert -Category 'ESXi Alert - Hardware' -Body $alertBody | Out-Null
                    Write-Host $alertBody -ForegroundColor Red
                }
            }
            } catch {
            $object = New-Object -TypeName PSObject -Property @{
                VMHost = $VMHost
                SensorName = "NA"
                Status = "FailedToQuery"
                CurrentReading = "FailedToQuery"
            } | Select-Object VMHost, SensorName, CurrentReading, Status
            $OutArr += $object
            }
        Disconnect-ViServer -Server $VMHost -Force -Confirm:$false 
    }
}

if ($uploadLog -eq $true) {
    $LogPath = "$PSScriptRoot\ESXi-hardware-scan.csv" 
    $OutArr | export-csv $LogPath  -Append -Force
    Upload-File -FilePath $LogPath | Out-Null
}
