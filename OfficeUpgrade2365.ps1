Start-Transcript -Path c:\temp\log.log

#Warunek sprawdzający popprawną wersję Office, czy ma 365, jeśli nie to zakończ działanie.

#detect Office if not 365 run this code
$BitLockerStatus = Get-BitLockerVolume | Where-Object {$_.VolumeType -like "OperatingSystem"}
if($BitLockerStatus.VolumeStatus -like "FullyEncrypted") {
    if($BitLockerStatus.ProtectionStatus -like "On") {
        $BitLockerStatus | Suspend-BitLocker -RebootCount 1
    }
}


#uninstall Office

#$encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($scriptAfterRestart.ToString()))
$action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument "-File c:\temp\OfficeUpgrade2365Install.ps1" 
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "Office365Installation" -Trigger $trigger -Action $action
Stop-Transcript
Restart-Computer -Force