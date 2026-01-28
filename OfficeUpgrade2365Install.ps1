Start-Transcript -Path c:\temp\logafterrestart.log
$BitLockerStatus = Get-BitLockerVolume | Where-Object {$_.VolumeType -like "OperatingSystem"}
if($BitLockerStatus.VolumeStatus -like "FullyEncrypted") {
    if($BitLockerStatus.ProtectionStatus -like "Off") {
        $BitLockerStatus | Resume-BitLocker
    }
}
#install Office
cmd /c msg /time=600 * "Aplikacja Office 365 została zainstalowana!"    
Get-ScheduledTask -TaskName Office365Installation -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false
Stop-Transcript