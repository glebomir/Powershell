$ErrorActionPreference = "Stop"
$path = Get-Content -Path E:\Scripts\variableList.json | ConvertFrom-Json

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

}

function download-7zip {
    
}

function download-PemHeart {

}