#region VERIFICA FILE appsettings.json
$FilePathAppsetting = "$($PSScriptRoot)/Config/appsettings.json"
try { 
    $Global:Appsetting = Get-Content -Raw -Path $FilePathAppsetting -ErrorAction Stop  | ConvertFrom-Json
}        
catch {
    $errorAppsettings = "ERROR: appsettings.json not found, empty or bad format file"
    Write-FatalLog $errorAppsettings
    New-Item -ItemType "file" -Path "$($PSScriptRoot)/Logs/" -Name "FATAL_LOG.txt" -Value $errorAppsettings -Force | Out-Null
    exit $LASTEXITCODE
}
#endregion

#region FUNCTIONS INVIO NOTIFICHE
<#
    .Description
    Invio delle notifiche tramite Send Grid
    $appsetting = parametri in ingressi da appsettings.json
    $Body = corpo della mail
#>
function Send-Mail() {
    Param (
        $appsetting,
        $Body
    )

$Parameters = @{
    FromAddress = $appsetting.EmailSettings.FromEmail
    ToAddress   = $appsetting.EmailSettings.ToEmail
    Subject     = $appsetting.EmailSettings.EmailSubject
    Body        = $Body
    Token       = $appsetting.EmailSettings.Token
    FromName    = $appsetting.EmailSettings.FromName
    AttachmentPath = $fileErrorList.FullName
}
Send-PSSendGridMail @Parameters
}

<#
    .Description
    Invio messaggi di Teams tramite WebHook
    $TeamSettings = parametri in ingressi da appsettings.json
    $Message = scritta che compare nel body del mesaagio
#>
function Send-TeamsMessage () {
    param(
      $TeamSettings,
      $Message
      
    )
    $JSONBody = [PSCustomObject][Ordered]@{
        "@type"      = "MessageCard"
        "@context"   = "http://schema.org/extensions"
        "summary"    = "Error Notifier SharePointBCK"
        "themeColor" = '0078D7'
        "title"      = "Details Below"
        "text"       = $Message
    }
    
    $TeamMessageBody = ConvertTo-Json $JSONBody -Depth 100
    
    $parameters = @{
        "URI"         = $TeamSettings.URI
        "Method"      = 'POST'
        "Body"        = $TeamMessageBody
        "ContentType" = 'application/json'
    }    
    $esitoTeams=Invoke-RestMethod @parameters
    if ($esitoTeams){
        Write-InfoLog "Notify Teams send"
    }else {
        Write-ErrorLog "ERROR: notify Teams not send"
    }
}
#endregion

#region INIZIALIZZA SERILOG
Import-Module PoShLog
Import-Module PoShLog.Enrichers

$ConsoleParams = @{ 
    OutputTemplate=$Appsetting.Serilog.ConsoleParams.outputTemplate
}

$FileParams= @{
    OutputTemplate=$Appsetting.Serilog.FileParams.OutputTemplate
    Path="$($PSScriptRoot)$($Appsetting.Serilog.FileParams.Path)"
    Formatter=Get-JsonFormatter
    Rollinginterval=$Appsetting.Serilog.FileParams.rollinginterval
    RetainedFileCountLimit=$Appsetting.Serilog.FileParams.RetainedFileCountLimit
    RestrictedToMinimumLevel=$Appsetting.Serilog.FileParams.RestrictedToMinimumLevel
 }

$minimumLevel=$Appsetting.Serilog.MinimumLevel

New-Logger |
    Add-EnrichWithEnvironment |
    Add-EnrichWithExceptionDetails |
    Set-MinimumLevel -Value $minimumLevel | 
    Add-SinkFile @FileParams |
    Add-SinkConsole @ConsoleParams|
    Start-Logger
#endregion

#region VERIFICA FILE sites.txt
$FileSites = "$($PSScriptRoot)/Config/sites.txt"
try {
    $Global:Sites = Get-Content $FileSites -ErrorAction Stop
}
catch {
    $errorSites = "ERROR: sites.txt not found or empty"
    Write-FatalLog $errorSites
    if ($useTeams) {
        Send-TeamsMessage -TeamSettings $TeamsSetting -Message $errorSites
    }
    if ($useEmail) {
        Send-Mail -appsetting $Appsetting -Body $errorSites
    }   
    exit $LASTEXITCODE
}
#endregion

#region CARICAMENTO VARIABILI
$Global:errorList = [System.Collections.ArrayList]@()
$Global:fileErrorList = @{
    FileName = "ErrorList.txt"
    Path = "$($PSScriptRoot)/Logs/"
    FullName = "$($PSScriptRoot)/Logs/ErrorList.txt"
}
New-Item -ItemType "file" -Path $fileErrorList.Path -Name $fileErrorList.FileName -Value "ERROR DETAILS: `r `n" -Force | Out-Null

$Global:OutputPath = "$($Appsetting.FolderSettings.rootBackup)$($Appsetting.FolderSettings.nameBackup)$($Date)"
$Global:tenant = $Appsetting.UserSettings.tenant
$Global:Longpath
$Global:replace

if ($storeInCloud) {
    $Global:date = Get-Date -Format "yyyy-MM-dd-hh-mm-ss"
    $Global:BucketName = "$($Appsetting.CloudSettings.bucketName)/$($date)"
}

if ($Appsetting.FolderSettings.OS -eq "Win") {
    $Longpath = "\\?\"
    $replace = .Replace("/", "\")
}

$Global:useTeams = $Appsetting.TeamsSetting.useTeams
$Global:useEmail = $Appsetting.EmailSettings.useEmail

if ($useTeams) {
    $Global:TeamsSetting = $Appsetting.TeamsSetting
}
#endregion

#region FUNCTION CONNESSIONE
<#
    .Description
    Connessione a Sharepoint online
#>
function ConnectToSPO{
	$AdminURL = "https://$tenant-admin.sharepoint.com"

    try {
        Connect-PnPOnline -Url $AdminURL -Interactive
    }
    catch {
        $errorConnection = $_.Exception.Message
        Write-ErrorLog $errorConnection
        $errorList.Add($errorConnection)
        if ($useTeams) {
            Send-TeamsMessage -TeamSettings $TeamsSetting -Message $errorConnection
        }
        if ($useEmail) {
            Send-Mail -appsetting $Appsetting -Body $errorConnection
        }        
        $errorList | Out-File -Append $fileErrorList.FullName
        exit $LASTEXITCODE
    }
}
#endregion

#region FUNCTION UPLOAD & DELETE ZIP
<#
    .Description
    Uplaod della cartella zip e cancellazione della medesima per risparmiare spazio sul primo disco
    $File = path del file 
#>
function UploadAndDelete {
    param (
        [string]$File,
        $BucketName
    )

    $FileName = $File.Substring(($File.lastindexof("\") + 1))
    $PathFile = $File
    $accesKey = $Appsetting.CloudSettings.accesKey
    $secretKey = $Appsetting.CloudSettings.secretKey
    $profileName = $Appsetting.CloudSettings.profileName
    $endpointUrl = $Appsetting.CloudSettings.endpointUrl
    $region = $Appsetting.CloudSettings.region
    
    try {
    # ------------ Set account user Wasabi ------------
    Set-AWSCredentials -AccessKey $accesKey -SecretKey $secretKey -StoreAs $profileName            
    }
    catch {
        $Global:errorUploadAndDelete = $_.Exception.Message
        Write-ErrorLog $errorUploadAndDelete
        $errorList.Add($errorUploadAndDelete)
    }

    # ------------ Upload cartella zip ------------
    Write-S3Object -BucketName $BucketName -Key $FileName -File $PathFile -EndpointUrl $endpointUrl -Region $region

    # ------------ Eliminazione zip ------------
    Write-InfoLog "Start delete:$($compress.DestinationPath)"
    Remove-Item $compress.DestinationPath -Recurse
    Write-InfoLog "End delete:$($compress.DestinationPath)"
}
#endregion

#region FUNCTION DOWNLOAD LIST FILES
<#
    .Description
    Download dei file di una singola raccolta, creazione dello zip del sito, upload del file zip e cancellazione del medesimo
    $files = file da scaricare
#>
function DownloadLibrary {
    param (
        $files
    )

    $result="OK"
    if ($files)
    {
        foreach ($file in $files)
        {                  
            $fullpath = "/" + ($file[0].Path.Identity -split (":file:/"))[1]
            Write-InfoLog "Downloading File: $($fullpath)"
            $pathDirectory = "$Longpat$($OutputPath)$($fullpath.ToString().Substring(0, (($fullpath.lastindexof("/") + 1))))$replace"
            try { $newdir = New-Item -ItemType "directory" -Path $pathDirectory -Force }
            catch {
                $Message = $_.Exception.Message
                $result=$Message
                Write-ErrorLog $Message
                $errorList.Add($Message)
                
            }
            $pathPnp = "$Longpat$($OutputPath)$($fullpath.ToString().Substring(0, (($fullpath.lastindexof("/") + 1))))$replace"
            $GetPnpFileParams = @{ }
            $GetPnpFileParams.Add('Url', $file.ServerRelativeUrl)
            $GetPnpFileParams.Add('Path', $pathPnp )
            $GetPnpFileParams.Add('Filename', $fullpath.Substring((($fullpath.lastindexof("/") + 1))))
            $GetPnpFileParams.Add('AsFile', $true)
            $GetPnpFileParams.Add('Force',$true)
            try {
                Get-PnpFile @GetPnpFileParams -ErrorAction Stop
            }
            catch {
                $errorPnp = $_.Exception
                $Message = "ERROR: $($file.ServerRelativeUrl) non copiato"
                Write-ErrorLog $Message
                Write-ErrorLog $errorPnp
                $result=$errorPnp
                $errorList.Add($errorPnp)
                Add-Content -Path $fileErrorList.FullName $Message
            }            
        }
        $errorList | Out-File -Append $fileErrorList.FullName
        Write-InfoLog "End downloading files from site $($Site)"
        return $result
    }
}
#endregion

#region DOWNLOAD DEI SITI
ConnectToSPO
$xlfile = "$($PSScriptRoot)/Logs/SiteReport.xlsx"
$rowExcel=1
Remove-Item $xlfile -ErrorAction SilentlyContinue
$Global:storeInCloud = $appsetting.CloudSettings.storeInCloud

if ($storeInCloud) {
    Import-Module AWSPowerShell
}

foreach ($item in $Sites) {

    $Site = "https://$tenant.sharepoint.com/sites/$item"
    $TestataSito = [PSCustomObject]@{
        Istanza  = $tenant
        Sito = $item
        url  = $Site
    }

    Export-Excel $xlfile -InputObject $TestataSito -AutoSize -StartRow $rowExcel -TableName $item
    $rowExcel+=2

    $cmd = "Connect-PnPOnline -Url $Site -Interactive"
    Invoke-Expression $cmd

    $DocumentLibraries = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false}
    $DocumentLibraries = $DocumentLibraries | Select-Object Title, Id, RootFolder
    $EsitiRaccolte=New-Object System.Collections.Generic.List[PSObject]
    foreach ($Library in $DocumentLibraries) {
        
        $FindParams = @{ }
        $FindParams.Add("Match", "*")        
        $FindParams.Add("List", $Library.Id)

        try { $global:files = Find-PnpFile @FindParams -ea stop}
        catch
        {
            $Message = $_.Exception.Message
            If ($Message -like "*object reference not set*")
            {
                $error1 = "Looks like the List $($Library.Title) or Folder $($Folder) isn't found in site $($Site)."
                Write-ErrorLog "Looks like the List $($Library.Title) or Folder $($Folder) isn't found in site $($Site)."
                $errorList.Add($error1)
            }
            Else
            {
                Write-ErrorLog $Message
                $errorList.Add($Message)
            }
            $s++
            continue # go to next folder            
        }

        Write-InfoLog "Start downloading files from site $($Site)" 
        $esitoDownLibrary = DownloadLibrary -files $files
        $pathLibrary="$($OutputPath)$($Library.RootFolder.ServerRelativeUrl)".Replace("/", "\")
        
        $EsitiRaccolte.Add([PSCustomObject]@{
            Title=$Library.Title
            Date=Get-Date -Format "dddd MM/dd/yyyy HH:mm"
            Esito=if($esitoDownLibrary -eq "OK"){"OK"}else{"FAIL"}
            Messaggio=$esitoDownLibrary.Message
        })

        
        if ($esitoDownLibrary -eq "OK") {
            Write-InfoLog "Download library:$($pathLibrary) positive"
        }else {
            Write-ErrorLog "Download library:$($pathLibrary) negative"
        }

        $compress = @{
            LiteralPath= "$($pathLibrary)"
            CompressionLevel = "Fastest"
            DestinationPath = "$($pathLibrary).zip"
        }    

        # ------------ Creazione zip della cartella temporanea ------------
        Write-InfoLog "Start creation:$($compress.DestinationPath)"
        Compress-Archive @compress -Force
        Write-InfoLog "End creation:$($compress.DestinationPath)"

        # ------------ Eliminazione cartella temporanea ------------
        Write-InfoLog "Start delete:$($compress.LiteralPath)"        
        Remove-Item $compress.LiteralPath -Recurse
        Write-InfoLog "End delete:$($compress.LiteralPath)"
        
        # ------------ Inizio upload e cancellazione libreria.zip ------------
        if ($storeInCloud) {
            try {
                
                Write-InfoLog "Start upload into Wasabi: $($Library.Title)"
                UploadAndDelete -File $compress.DestinationPath -BucketName "$BucketName/$item"
                Write-InfoLog "End upload into Wasabi: $($Library.Title) and delete .zip"
            }
            catch {
                $Message = $_.Exception.Message
                Write-ErrorLog $Message
                $errorList.Add($Message)
            }  
        }                   
 
    }  

    # ------------ Esportazione esiti raccolte in Excel ------------
    Export-Excel $xlfile -InputObject $EsitiRaccolte -AutoSize -StartRow $rowExcel -TableName "$($item)detail"
    $rowExcel+=($EsitiRaccolte.Count+2)
}

Write-InfoLog "Excel created"
#endregion

$errorList | Out-File -Append $fileErrorList.FullName

if (errorList) {
    $result = ("ERROR: Backup Complete with errors")
}else {
    $result = ("Backup Complete")
}

Write-InfoLog $result

if ($useTeams) {
    Send-TeamsMessage -TeamSettings $TeamsSetting -Message $result
}

if ($useEmail) {
    Send-Mail -appsetting $Appsetting -Body $result -logMail $fileErrorList.FullName
}
