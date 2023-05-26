$ErrorActionPreference = "Stop"
function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
}
#$path = Get-ScriptDirectory #uruchamiać w trybie nienadzorowanych.
$path = Get-Content -Path C:\scripts\DownloadAppsNew\variableList.json | ConvertFrom-Json

$logFile = "NewSoftwareList.csv"
#wstaw funkcje  ustawiajaca katalog
function ClearLogs {
    param (
        [PSCustomObject] $path
    )

    Remove-Item -Path $($path.LogLocation + "\$logFile") -Force -ErrorAction SilentlyContinue
}

function download-Firefox { 
    param (
        [PSCustomObject] $path
    )
    $latestVersion = (Invoke-WebRequest  "https://product-details.mozilla.org/1.0/firefox_versions.json" | ConvertFrom-Json).LATEST_FIREFOX_VERSION
    $downloadUri = "https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=pl"
    
    if(!$(Get-ChildItem -Path $($path.FirefoxOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.FirefoxOutputLocation + "\Mozilla Firefox $latestVersion" +".msi")
            "Mozilla Firefox;$latestVersion; Pobrane" | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "Mozilla Firefox;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }
    elseif($(Get-ChildItem -Path $($path.FirefoxOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)  -notlike "*$latestVersion*"){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.FirefoxOutputLocation + "\Mozilla Firefox $latestVersion" +".msi")
            "Mozilla Firefox;$latestVersion; Pobrane" | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "Mozilla Firefox;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }
}

function download-Chrome {
    param (
        [PSCustomObject] $path
    )
    $latestVersion = (((Invoke-WebRequest "https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions/all/releases?filter=endtime=none").Content | ConvertFrom-Json).releases | Sort-Object -Property fraction -Descending | Select -First 1).Version
    $downloadUri = "https://dl.google.com/chrome/install/ChromeStandaloneSetup64.msi"
    
    if(!$(Get-ChildItem -Path $($path.ChromeOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.ChromeOutputLocation + "\Google Chrome $latestVersion" +".msi")
            "Google Chrome;$latestVersion; Pobrane" | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "Google Chrome;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }
    elseif($(Get-ChildItem -Path $($path.ChromeOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)  -notlike "*$latestVersion*"){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.ChromeOutputLocation + "\Google Chrome $latestVersion" +".msi")
            "Google Chrome;$latestVersion; Pobrane" | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "Google Chrome;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }
}
    
function download-LibreOffice {
    param (
        [PSCustomObject] $path
    )
    $downloadUri = ((invoke-webrequest -uri "https://www.libreoffice.org/download/download-libreoffice/?type=win-x86_64&version=7.5.3&lang=pl").links | ForEach-Object {
        $_ | Where-Object {$_.href -like "*_Win_x86-64.msi"}
    }).href
    $latestVersion = ($downloadUri.Replace("https://www.libreoffice.org/donate/dl/win-x86_64/","")) -replace '/.*',''
    
     if(!$(Get-ChildItem -Path $($path. + "\*") -Include "*.msi" | Sort-Object | Select -First 1)){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.LibreOfficeOutputLocation + "\LibreOffice $latestVersion" +".msi")
            "LibreOffice;$latestVersion; Pobrane" | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "LibreOffice;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }
    elseif($(Get-ChildItem -Path $($path.SevenZipOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)  -notlike "*$latestVersion*"){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.LibreOfficeOutputLocation + "\LibreOffice $latestVersion" +".msi")
            "LibreOffice;$latestVersion; Pobrane"  | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "LibreOffice;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }   
}

function download-7zip {
    param (
        [PSCustomObject] $path
    )
    $downloadUri = (invoke-webrequest -uri "https://7-zip.org.pl/sciagnij.html").links | ForEach-Object {
        $_ | Where-Object {$_.href -like "*-x64.msi*"}
    }
    $downloadUri = ($downloadUri | select -First 1).href
    $latestVersion = ($downloadUri.Replace("https://7-zip.org/a/7z","")).Replace("-x64.msi","")

    if(!$(Get-ChildItem -Path $($path.SevenZipOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.SevenZipOutputLocation + "\7Zip $latestVersion" +".msi")
            "7Zip;$latestVersion; Pobrane" | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "7Zip;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }
    elseif($(Get-ChildItem -Path $($path.SevenZipOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)  -notlike "*$latestVersion*"){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.SevenZipOutputLocation + "\7Zip $latestVersion" +".msi")
            "7Zip;$latestVersion; Pobrane"  | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "7Zip;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }   
}

function download-PemHeart {
    param (
        [PSCustomObject] $path
    )
    $downloadUri = ((invoke-webrequest -uri "https://www.cencert.pl/do-pobrania/oprogramowanie-do-podpisu/").links | ForEach-Object {
        $_ | Where-Object {($_.href -like "*.exe") -and ($_.href -like "*www.cencert.pl/*")}
    }).href
    $latestVersion = (($downloadUri.Replace("https://www.cencert.pl/wp-content/software/PH-",""))).Replace(".exe","")
    param (
        [PSCustomObject] $path
    )
    if(!$(Get-ChildItem -Path $($path.PemHeartOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.PemHeartOutputLocation + "\Pem-Heart $latestVersion" +".exe")
            "Pem-Heart;$latestVersion; Pobrane" | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "Pem-Heart;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    }
    elseif($(Get-ChildItem -Path $($path.SevenZipOutputLocation + "\*") -Include "*.msi" | Sort-Object | Select -First 1)  -notlike "*$latestVersion*"){
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path.PemHeartOutputLocation + "\Pem-Heart $latestVersion" +".exe")
            "Pem-Heart;$latestVersion; Pobrane"  | Add-Content $($path.LogLocation + "\$logFile")
        }
        catch {
            "Pem-Heart;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
    } 
}
