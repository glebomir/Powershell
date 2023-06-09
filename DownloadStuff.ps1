$ErrorActionPreference = "Stop"
function Get-ScriptDirectory
{
  $Invocation = (Get-Variable MyInvocation -Scope 1).Value
  Split-Path $Invocation.MyCommand.Path
}
$CurrentCatalog = Get-ScriptDirectory #uruchamiać w trybie nienadzorowanym.
#$CurrentCatalog = "C:\scripts\DownloadAppsNew"
$path = Get-Content -Path $CurrentCatalog\variableList.json | ConvertFrom-Json

$logFile = "NewSoftwareList.csv"
#wstaw funkcje  ustawiajaca katalog
function ClearLogs {
    param (
        [PSCustomObject] $path
    )
    $LogPath = $path.LogLocation
    Remove-Item -Path $($LogPath + "\$logFile") -Force -ErrorAction SilentlyContinue
}

function download-Soft {
     param (
        [string] $path,
        [string] $Name,
        [string] $logFile,
        [string] $ext,
        [string] $LogPath,
        [string] $downloadUri,
        [string] $latestVersion
     )
     if(!$(Test-Path -Path $($path + "\$Name $latestVersion" + "$ext"))) {
        try {
            Invoke-WebRequest -Uri $downloadUri -OutFile $($path + "\$Name $latestVersion" +"$ext")
            "$Name;$latestVersion; Pobrane" | Add-Content $($LogPath + "\$logFile")
        }
        catch {
            "Name;$latestVersion;Błąd pobierania" | Add-Content $($path.LogLocation + "\$logFile")
        }
     }     
}

function download-Firefox { 
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.FirefoxOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "Mozilla Firefox"
    $ext = ".msi"
    $latestVersion = (Invoke-WebRequest  "https://product-details.mozilla.org/1.0/firefox_versions.json" | ConvertFrom-Json).LATEST_FIREFOX_VERSION
    $downloadUri = "https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=pl"
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion    
}

function download-Chrome {
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.ChromeOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "Google Chrome"
    $ext = ".msi"
    $latestVersion = (((Invoke-WebRequest "https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions/all/releases?filter=endtime=none").Content | ConvertFrom-Json).releases | Sort-Object -Property fraction -Descending | Select -First 1).Version
    $downloadUri = "https://dl.google.com/tag/s/appname%3DGoogle%2520Chrome%26needsadmin%3Dtrue%26ap%3Dx64-stable-statsdef_0%26brand%3DGCEA/dl/chrome/install/googlechromestandaloneenterprise64.msi"
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion 
}
    
function download-LibreOffice {
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.LibreOfficeOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "Libre Office"
    $ext = ".msi"
    $downloadUri = ((invoke-webrequest -uri "https://www.libreoffice.org/download/download-libreoffice/?type=win-x86_64&lang=pl").links | ForEach-Object {
        $_ | Where-Object {$_.href -like "*_Win_x86-64.msi"}
    }).href
    $latestVersion = ($downloadUri.Replace("https://www.libreoffice.org/donate/dl/win-x86_64/","")) -replace '/.*',''
    $downloadUri = "https://download.documentfoundation.org/libreoffice/stable/" + $latestVersion +"/win/x86_64/LibreOffice_" + $latestVersion + "_Win_x86-64.msi"
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion
}

function download-7zip {
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.SevenZipOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "7Zip"
    $ext = ".msi"
    $downloadUri = (invoke-webrequest -uri "https://7-zip.org.pl/sciagnij.html").links | ForEach-Object {
        $_ | Where-Object {$_.href -like "*-x64.msi*"}
    }
    $downloadUri = ($downloadUri |select -Skip 1 | select -First 1).href
    $latestVersion = ($downloadUri.Replace("https://7-zip.org/a/7z","")).Replace("-x64.msi","")
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion   
}

function download-PemHeart {
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.PemHeartOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "Pem-Heart"
    $ext = ".exe"
    $downloadUri = ((invoke-webrequest -uri "https://www.cencert.pl/do-pobrania/oprogramowanie-do-podpisu/").links | ForEach-Object {
        $_ | Where-Object {($_.href -like "*.exe") -and ($_.href -like "*software*")}
    }).href | Select -First 1
    $downloadUri = "https://www.cencert.pl" + $downloadUri
    $latestVersion = (($downloadUri.Replace("https://www.cencert.pl/wp-content/software/PH-",""))).Replace(".exe","")
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion 
}

function download-AcrobatReaderDC {
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.AcrobatReaderDCOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "Acrobat Reader DC"
    $ext = ".msp"
    $latestVersion = (Invoke-RestMethod -Uri "https://rdc.adobe.io/reader/products?lang=en&site=enterprise&os=Windows 10&api_key=dc-get-adobereader-cdn").Products.Reader.Version
    $downloadUri = "https://ardownload2.adobe.com/pub/adobe/acrobat/win/AcrobatDC/" + $($latestVersion.Replace('.','')) + "/AcroRdrDCx64Upd"+ $($latestVersion.Replace('.','')) + $ext
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion
}

function download-WinSCP {
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.WinSCPOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "WinSCP"
    $ext = ".exe"
    #dzieli stringa po nowej linii   
    $downloadUri = (Invoke-RestMethod -Uri "https://winscp.net/eng/download.php").Split([System.Environment]::NewLine,[System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object {
        $_ | Where-Object { $_ -like "*Setup.exe*" }
    }
    $downloadUri = ("https://winscp.net" + $($downloadUri.Replace('<a href="','') -replace '" class.*','')).Replace(' ','')
    $latestVersion = ($downloadUri.Replace('https://winscp.net/download/WinSCP-','')).Replace('-Setup.exe','')
    $downloadUri = ((Invoke-WebRequest -Uri $downloadUri -UseBasicParsing).links | ForEach-Object { #jakiś skrypt otwiera na stronie uruchamia pobieranie dlatego ten dodatkowy zabieg
        $_ | Where-Object { $_.href -like "*$latestVersion-Setup.exe" }
    }).href
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion
}

function download-Szafir {
     param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $DownloadPath = $path.SzafirOutputLocation
    $LogPath = $path.LogLocation
    $AppName = "Szafir"
    $ext = ".msi"      
    $downloadUri = ((Invoke-WebRequest "https://www.elektronicznypodpis.pl/informacje/aplikacje/").links | ForEach-Object {
    $_ | Where-Object {$_ -like "*x64_jre11.msi*"}
    }).href
    $latestVersion = (($downloadUri.Replace('/gfx/elektronicznypodpis/userfiles/szafirsdk/szafir/instalator/szafir_','')).Replace('_x64_jre11.msi','')).Replace('_',' ')
    $downloadUri ="https://www.elektronicznypodpis.pl" + $downloadUri
    download-Soft -path $DownloadPath -Name $AppName -ext $ext -logFile $logFile -LogPath $LogPath -downloadUri $downloadUri -latestVersion $latestVersion
}

function Gather-Info {
    param (
        [PSCustomObject] $path,
        [string] $logFile
    )
    $LogPath = $path.LogLocation
    $newVersions = Import-Csv -Path $LogPath\$logFile -Delimiter ';' -Header Aplikacja,Wersja,Status  | ConvertTo-html
    
    $i = 0
    foreach($n in $newVersions){
    $i++
        if($n -like "</head><body>"){
            break;
        }
    }
    $nV = $newVersions[$i..$($newVersions.Length - 2)]
    $mailContent = @"
    <br></br>
    <h1>Nowe wersje aplikacji:</h1>
    <br></br>
    $nV
"@
    $mailContent
}

ClearLogs -path $path

download-Firefox -path $path -logFile $logFile
download-LibreOffice -path $path -logFile $logFile
download-7zip -path $path -logFile $logFile
download-PemHeart -path $path -logFile $logFile
download-Chrome -path $path -logFile $logFile
download-AcrobatReaderDC -path $path -logFile $logFile
download-WinSCP -path $path -logFile $logFile
download-Szafir -path $path -logFile $logFile

Import-Module "$($path.SendMailModulePath)\SendMail.psm1"
$VariableList =  Get-Content -Path "$CurrentCatalog\mailvariablelist.json"
$VariableList = $VariableList | ConvertFrom-Json
$LogPath = $path.LogLocation
if(Test-Path -Path "$LogPath\$logFile"){
    $MailBody = Gather-Info -path $path -logFile $logFile
    $body = Make-HTMLcontent -htmlContent $MailBody
    Get-Mail -args $VariableList -Body $body
    ClearLogs -path $path
}
