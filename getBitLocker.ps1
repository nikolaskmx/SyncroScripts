<#
.SYNOPSIS
    Checks for BitLocker 
.DESCRIPTION
    Scripts checks if BitLocker is enabled and uploads data to custom field "Bitlocker_Keys_C" in Syncro. If Syncro variable $enableBitLocker
    is set to $true and all prerequisites are met then it enables BitLocker on all drives.
    All information is uploaded in file to Syncro Asset for record keeping and Azure AD
    Required setup in Syncro:
        - variable (name: enableBitLocker) (type: dropdown with values: Yes,No)
        - asset custom field: BitLocker_Keys_C
.EXAMPLE
    Run script from Syncro as System 
.NOTES
    Author: Mariusz Sztanga
    Date:   May, 11, 2021   
#>
Import-Module $env:SyncroModule -WarningAction SilentlyContinue

if ($enableBitLocker -eq $null) {$enableBitLocker = "No"}

$TPMEnabled = (Get-TPM).TpmEnabled
$WindowsVer = Get-WmiObject -Query 'select * from Win32_OperatingSystem where (Version like "6.2%" or Version like "6.3%" or Version like "10.0%") and ProductType = "1"' -ErrorAction SilentlyContinue
$BitLockerReadyDrive = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction SilentlyContinue
$BitLockerDecrypted = Get-BitLockerVolume -MountPoint $env:SystemDrive | Where-Object {$_.VolumeStatus -eq "FullyDecrypted"} -ErrorAction SilentlyContinue
$BLVS = Get-BitLockerVolume | Where-Object {$_.KeyProtector | Where-Object {$_.KeyProtectorType -eq 'RecoveryPassword'}} -ErrorAction SilentlyContinue

#Step 1 - Check if TPM is enabled and initialise if required
if ($WindowsVer -and !$TPMEnabled) {
    Initialize-Tpm -AllowClear -AllowPhysicalPresence -ErrorAction SilentlyContinue
}

#Step 2 - Check if BitLocker volume is provisioned and partition system drive for BitLocker if required
if ($WindowsVer -and $TPMEnabled -and !$BitLockerReadyDrive) {
    Get-Service -Name defragsvc -ErrorAction SilentlyContinue | Set-Service -Status Running -ErrorAction SilentlyContinue
    BdeHdCfg -target $env:SystemDrive shrink -quiet
}

#Step 3 - If all prerequisites are met, then enable BitLocker
if ($WindowsVer -and $TPMEnabled -and $BitLockerReadyDrive -and $BitLockerDecrypted -and $enableBitLocker -eq "Yes") {
    Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -TpmProtector
    Enable-BitLocker -MountPoint $env:SystemDrive -RecoveryPasswordProtector -ErrorAction SilentlyContinue
    Log-Activity -Message "Enabled BitLocker" -EventName "BitLocker"
} else {
    Write-Host "Can't enable BitLocker - Windows: $WindowsVer TPM Enable: $TPMEnabled  BitLocker Drive Ready: $BitLockerReadyDrive BitLocker Decrypted: $BitLockerDecrypted  Enable: $enableBitLocker"
}

#Step 4 - Save info to file and upload C password to Syncro and AAD
if ($BLVS) {
    $fileValue = "BitLocker information for $Env:COMPUTERNAME"
    $fileValue += "------------------------------------------"
    ForEach ($BLV in $BLVS) {
        $AADbackup = BackupToAAD-BitLockerKeyProtector -MountPoint $env:SystemDrive -KeyProtectorId ((Get-BitLockerVolume -MountPoint $env:SystemDrive ).KeyProtector | Where-Object {$_.KeyProtectorType -eq "RecoveryPassword" }).KeyProtectorId
        if ($BLV.MountPoint -eq $env:SystemDrive) {
            Set-Asset-Field -Name "Bitlocker_Keys_C" -Value $($BLV.KeyProtector.RecoveryPassword)
        }
        $fileValue = @"
######### BitLocker information for $($BLV.MountPoint) #########
EncryptionMethod     : $($BLV.EncryptionMethod)
AutoUnlockEnabled    : $($BLV.AutoUnlockEnabled)
AutoUnlockKeyStored  : $($BLV.AutoUnlockKeyStored)
MetadataVersion      : $($BLV.MetadataVersion)
VolumeStatus         : $($BLV.VolumeStatus)
ProtectionStatus     : $($BLV.ProtectionStatus)
LockStatus           : $($BLV.LockStatus)
EncryptionPercentage : $($BLV.EncryptionPercentage)
WipePercentage       : $($BLV.WipePercentage)
VolumeType           : $($BLV.VolumeType)
CapacityGB           : $($BLV.CapacityGB)
KeyProtector         : $($BLV.KeyProtector)

RecoveryPassword     : $($BLV.KeyProtector.RecoveryPassword)

Drives backup to AAD : $AADbackup

"@
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $filePath = "$PSScriptRoot\BitLocker-"+$timestamp+".log"
    Add-Content -Path $filePath -Value $fileValue

    # Upload file and wait 30 sec before deleting 
    Upload-File -FilePath $filePath
    Start-Sleep 30
    Remove-Item $filePath
}
    
Exit 0