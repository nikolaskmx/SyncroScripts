<#
.SYNOPSIS
    This script will add option to send Syncro Messages to users when running scripts as SYSTEM.
    
.DESCRIPTION
    This is missing functionality from Syncro, sending message to user(s) when running as SYSTEM
    It will display default Syncro popup with your logo to all logged in users. 
    sDelay option gives user delay in seconds from the time you send message to when it displays
.EXAMPLE
    MessageAsSystem -Title "Test" -Message "Test message" 
    MessageAsSystem -Title "Test" -Message "Test message" -sDelay 120
.NOTES
    Author: Mariusz Sztanga
    Date:   April, 30, 2021   
#>
function MessageAsSystem {
    Param
    ( 
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [int]$sDelay = 15
    )
    $when = (Get-Date).AddSeconds($sDelay) | Get-Date -UFormat "%T"
    $argument = "--broadcast-message `"$Message`" --broadcast-title `"$Title`""
    $dUN = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='explorer.exe'" -ErrorAction SilentlyContinue).GetOwner()
    if ($dUN.Count -ne 0) {
        Foreach ($cUN in $dUN) {
            $username = ($cUN.Domain)+"\"+($cUN.User)
            $action = New-ScheduledTaskAction -Execute "$env:ProgramFiles\RepairTech\Syncro\Syncro.App.Runner.exe" -Argument $argument
            $trigger = New-ScheduledTaskTrigger -Once -At $when
            $principal = New-ScheduledTaskPrincipal -UserId $username
            $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal
            Register-ScheduledTask "BroadcastMessage-$username" -InputObject $task | Out-Null
            Start-Sleep -Seconds (30+$sDelay)
            Unregister-ScheduledTask -TaskName "BroadcastMessage-$username" -Confirm:$false
        }
    } else {
        Write-Host "Interactive User not DETECTED"
    }
}