param(
    [string]$Path="$PWD\MpSupportFiles.cab",
    [string]$Result,
    [switch]$FullEventLog,
    [string]$MpLogFolder
)
 
# Analyze MpSupportFiles.cab to get a result of signature update issue
$MpCabFolder = Split-Path $Path
$LogFileName = (Split-Path $Path -Leaf).split(".")[0]
if(!$MpLogFolder) { $MpLogFolder = $MpCabFolder + "\" + $LogFileName } 
$AnalysisLog = "$MpLogFolder\AnalysisLog.txt"
if(!$Result) { $Result="$MpLogFolder\SignatureUpdateAnalysis.txt" }
# $FALLBACK_SOURCES = @("InternalDefinitionUpdateServer", "MicrosoftUpdateServer", "MMPC", "FileShares")

function List-StartSection([string]$Value){
    Write-Host -ForegroundColor Cyan "[Start] $Value"
}

function List-EndSection([string]$Value){
    Write-Host -ForegroundColor Yellow "[End] $Value"
}

function List-Info($Name, $Value){
    Write-Host -ForegroundColor Green "$Name : " -NoNewline
    Write-Host -ForegroundColor Yellow "$Value"
    "$Name : $Value" | Out-File $Result -Append
}

function List-Warning($Value){
    Write-Host -ForegroundColor Red "[!] $Value"
    "[!] $Value" | Out-File $Result -Append
}

function List-Success($Value){
    Write-Host -ForegroundColor Green "[+] $Value"
    "[+] $Value" | Out-File $Result -Append
}

function Convert-RegDateTime([string]$date_string){
    # Sample string date like this: "[UTC] Sunday, August 11, 2024 8:55:06 AM"
    $format = "dddd, MMMM d, yyyy h:mm:ss tt"
    $dateString_Trimmed = $date_string -replace '^\[UTC\] ', ''
    $dateTime = [DateTime]::ParseExact($dateString_Trimmed, $format, [System.Globalization.CultureInfo]::InvariantCulture)
    $dateTimeUtc = [DateTime]::SpecifyKind($dateTime, [System.DateTimeKind]::Utc)

    return $dateTimeUtc
}

function Get-LastEventsById($EventLog, $Id) {
    if (!($EventLog -Like "*.csv")){ Write-Host "Wrong File extension"; return 1}
    $FilteredEvents = Import-Csv $EventLog | Where-Object EventId -eq $Id

    if ($FilteredEvents.Count -gt 0) { return $FilteredEvents[0] }
    else { return 1 }
}

$UserDecision = ""
if (Test-Path $Result){
    while ($OverWrite -ne $true){
        switch ($UserDecision){
            'y'{
                Write-Host "Overwriting $Result"
                $now = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
                Write-Output "Log analyzed at $now" | Set-Content $Result -Force
                $OverWrite = $true
            }
            'n'{
                Write-Host -ForegroundColor Yellow "List existing results."
                Get-Content $Result
                Write-Host -ForegroundColor Yellow "Exit."
                return
            }
            Default {
                Write-Host -ForegroundColor Cyan "Result file already exists: $Result"
                $UserDecision = Read-Host "Do you want to overwrite it? (y/n)"
            }
        }
    }
}

# Extract Defender logs
List-StartSection "Extract Defender logs"
if (!(Test-Path $Path)){
    Write-Host "File does not exist: $Path"
}
if (!(Test-Path $MpLogFolder)){ 
    New-Item -Name $LogFileName -ItemType Directory -Path $MpCabFolder | Out-Null
}
# Expand MpSupportFiles.cab
# Source: C:\Windows\system32\expand.exe
if (!(Test-Path $AnalysisLog)){
    Write-Host "[*] Extracting CAB file using expand to folder: $MpLogFolder"
    expand -i -f:* $Path $MpLogFolder | Out-Null
    New-Item -Name AnalysisLog.txt -ItemType File -Path $MpLogFolder | Out-Null
    Write-Host "[+] Extract .cab file to $MpLogFolder Done." | Out-File $AnalysisLog -Force
}
else {
    Write-Host "[+] 'AnalysisLog.txt' already exists. Skip extraction."
}

List-EndSection "Extraction Defender logs ends."

# Organize the files
$DynamicSignatureFiles = Get-ChildItem $MpLogFolder | Where-Object Name -Match "[a-f0-9]{40}"
# $WindowsUpdateEtl = Get-ChildItem $MpLogFolder | Where-Object Name -Like WindowsUpdate*.etl
New-Item -ItemType Directory -Path "$MpLogFolder\WindowsUpdateEtlLog", "$MpLogFolder\DynamicSignatures" 2>&1 | Out-Null
Move-Item -Path $MpLogFolder\WindowsUpdate*.etl -Destination $MpLogFolder\WindowsUpdateEtlLog -Force
if ($DynamicSignatureFiles.Length -gt 0) { Move-Item -Path $DynamicSignatureFiles -Destination $MpLogFolder\DynamicSignatures -Force}

# Convert MpRegistry.txt to JSON.
Convert-MpRegistrytxtToJson -Path "$MpLogFolder\MpRegistry.txt" -OutFile "$MpLogFolder\MpRegistry.json"
Write-Host "[*] Convert MpCmdRun-System.log. It may takes long time if the log size is big."
if (Test-Path "$MpLogFolder\MpCmdRun-SystemTemp.log"){
    Rename-Item -Path "$MpLogFolder\MpCmdRun-SystemTemp.log" -NewName MpCmdRun-System.log
}
if (!(Test-Path "$MpLogFolder\MpCmdRun-system.json")){
    Read-MpCmdRunLog -logFilePath "$MpLogFolder\MpCmdRun-System.log" -OutFile "$MpLogFolder\MpCmdRun-System.json" -quiet *>&1 | Out-Null
}
Write-Host "Done."

# Windows Update log
List-StartSection "Convert Windows Update ETL file."
if (Test-Path "$MpLogFolder\WindowsUpdate.log") { Write-Host "[+] 'WindowsUpdate.log' already exists." }
else { Get-WindowsUpdateLog -ETLPath "$MpLogFolder\WindowsUpdateEtlLog" -LogPath "$MpLogFolder\WindowsUpdate.log" *>&1 | Out-Null } # default output: WindowsUpdate.log
if (Test-Path "$MpLogFolder\WindowsUpdate.log") { Write-Host "[+] Done." }
else { Write-Host -ForegroundColor "[-] Failed to use Get-WindowsUpdateLog to convert WindowsUpdate.log" }

# Converting Logs
List-StartSection "Convert Defender logs."
# Convert WindowsUpdate.log to csv
if (!(Test-Path "$MpLogFolder\WindowsUpdateLog.csv")){
    Read-WindowsUpdateLog -Path "$MpLogFolder\WindowsUpdate.log" -OutFile "$MpLogFolder\WindowsUpdateLog.csv"
}
else { Write-Host "[+] 'WindowsUpdate.csv' already exists." }

# Merge and convert MpOperationalEvents.txt.bak
$EventLogs = "$MpLogFolder\MpOperationalEvents.txt"
if (Test-Path "$EventLogs.bak"){
    Get-Content "$EventLogs.bak", "$EventLogs" -Encoding unicode | Set-Content "$MpLogFolder\MpOperationalEvents-merged.txt" -Force -Encoding unicode
    $EventLogs = "$MpLogFolder\MpOperationalEvents-merged.txt"
}

if (!$FullEventLog) {
    Get-Content $EventLogs -Encoding unicode -head 30000 | Set-Content "$MpLogFolder\MpOperationalEvents-head30000.txt" -Force -Encoding unicode
    Write-Host "[*] 'FullEventLog' switch is not used. Analyze only first 30000 lines of the event log for saving time."
    $EventLogs = "$MpLogFolder\MpOperationalEvents-head30000.txt"
}

# Convert Event logs to csv or json.
Write-Host "[*] Convert MpOperationalEventLog. It might take some time if the log is too long..."
if (!(Test-Path "$MpLogFolder\MpOperationalEvents.csv")) {Convert-MpOperationalEventLogTxt -Path $EventLogs -OutFile "$MpLogFolder\MpOperationalEvents.csv" 2>&1 3>&1 | Out-Null}
else {Write-Host "[+] MpOperationalEvents.csv already exists." }
List-EndSection "Convert Logs."

# Main
# Get Signature information and fallback order
$DefenderConfigs = Get-Content "$MpLogFolder\MpRegistry.json" | ConvertFrom-Json
$SignatureConfigs = $DefenderConfigs."effective policy"."Signature Updates"
$SharedSignatureRoot = $SignatureConfigs."SharedSignatureRoot"
$FallbackOrder = $SignatureConfigs."FallbackOrder"
$AuGracePeriod = $SignatureConfigs."AuGracePeriod".Split()[0]
$SignatureVersion = $SignatureConfigs.'AVSignatureVersion'
# Timestamps
$AVSignatureApplied = Convert-RegDateTime $SignatureConfigs."AVSignatureApplied" # Represents the dattime stamp when AV Signature version was created
$SignaturesLastUpdated = Convert-RegDateTime $SignatureConfigs."SignaturesLastUpdated"
$LastFallbackTime = Convert-RegDateTime $SignatureConfigs."LastFallbackTime"

# List the most interested information
Write-Host "`n"
List-StartSection "List some interested information."
Write-Host "`n"
List-Info "Signature version" $SignatureVersion 
List-Info "Signature created" $AVSignatureApplied 
List-Info "Signature was last updated" $SignaturesLastUpdated 
List-Info "Fallback order" $FallbackOrder 
List-Info "Last Fallback Time" $LastFallbackTime 
Write-Host "`n"

# List last update success (2000) and failure (2001) events.
$LastUpdateErrorEvent = Get-LastEventsById "$MpLogFolder\MpOperationalEvents.csv" "2001"
$LastUpdateSuccessEvent = Get-LastEventsById "$MpLogFolder\MpOperationalEvents.csv" "2000"
if ($LastUpdateErrorEvent -ne 1) { 
    Write-Host -ForegroundColor Cyan "Lastest Event of Update failure:"
    $LastUpdateErrorEvent | Format-List
    $LastUpdateErrorEvent | Out-File $Result -Append
}
if ($LastUpdateSuccessEvent -ne 1){
    Write-Host -ForegroundColor Cyan "Lastest Event of Update success:"
    $LastUpdateSuccessEvent | Format-List
    $LastUpdateSuccessEvent | Out-File $Result -Append
}

# Start analyzing.
List-StartSection "Start analyzing signature update."
# Check SharedSignatureRoot.
if ($SharedSignatureRoot -ne ""){
    List-Warning "'SharedSignatureRoot' is not empty. Fallback order is ignored." 
    List-Warning "No signature update will be triggered when SharedSignatureRoot is $SharedSignatureRoot" 
    Write-Host "Exit."
    return
}
else {
    Write-Host "[+] 'SharedSignatureRoot' is empty. Continue." 
}

# Check SCCM related
if ($AuGracePeriod -ne "0"){
    Write-Host "[!] SCCM is managing signature update. AuGracePeriod is $AuGracePeriod minutes." 
    $AuGraceExpiration = $SignaturesLastUpdated.AddMinutes($AuGracePeriod)
    if ($AuGraceExpiration -gt $LastFallbackTime) {
        Write-Host "Grace Period is not expired. Defender is being update by SCCM and Fallback order will not be followed." 
        Write-Host "Please confirm with the customer or collaborate with SCCM team." 
        return
    }
}
elseif ($AuGracePeriod -eq "0"){
    Write-Host "[+] AuGracePeriod is 0. SCCM is not managing signature update. Fallback Order will be analyzed" 
}

# Analyze fallback order.
$Fallbackorder -replace "[{}]","" | Out-Null
$UpdateSources = $Fallbackorder.split('|')
$SourceOrder = 0
$WUEvents = Import-Csv "$MpLogFolder\WindowsUpdateLog.csv"
$WuConfigs = $DefenderConfigs."WindowsUpdate"."HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\WindowsUpdate"
if($WuConfigs."DoNotConnectToWindowsUpdateInternetLocations"){$DoNotConnectToWindowsUpdateInternetLocations = $WuConfigs."DoNotConnectToWindowsUpdateInternetLocations".split()[0]}
if($WuConfigs."SetPolicyDrivenUpdateSourceForOtherUpdates"){$SetPolicyDrivenUpdateSourceForOtherUpdates = $WuConfigs."SetPolicyDrivenUpdateSourceForOtherUpdates".split()[0]}
$ForceWSUS = $false
if($WuConfigs."WUServer"){$WUServer = $WuConfigs."WUServer"}

$FallbackOrderResults = [PSCustomObject]@{}

foreach ($UpdateSource in $UpdateSources){
    switch ($UpdateSource) {
        'MicrosoftUpdateServer' { 
            # MU/WU source
            List-StartSection "[*] Analyzing MU/WU source" 

            # Check if MU/WU is updateing from WSUS
            if ($DoNotConnectToWindowsUpdateInternetLocations -eq "1"){
                Write-Host "[!] MU/WU cannot connect to Internet location due to 'DoNotConnectToWindowsUpdateInternetLocations' is set to $DoNotConnectToWindowsUpdateInternetLocations" 
                $ForceWSUS = $true
            }
            elseif ($SetPolicyDrivenUpdateSourceForOtherUpdates -eq "1"){
                Write-Host "[!] MU/WU cannot connect to Internet location due to 'SetPolicyDrivenUpdateSourceForOtherUpdates' is set to $SetPolicyDrivenUpdateSourceForOtherUpdates" 
                Write-Host "[!] Most likely, SCCM is managing the updates, and forcing updates from SCCM/WSUS." 
                $ForceWSUS = $true
            }

            if(!$WUServer -and $ForceWSUS){ 
                Write-Host "[-] MU/WU is forced to use WSUS, but WUServer is not configured." | Out-File $Result -Append
                List-Warning "Cannot update from MU/WU or WSUS. Exit." 
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.WUServer" -Value "[Error] Force WSUS and empty WSUS."
                continue
            }
            elseif($WUServer){
                Write-Host "[+] MU/WU can download update from WSUS. WUServer: $WUServer" | Out-File $Result -Append

                # Find search result
                $LastAgentInitiateTime = [datetime](($WUEvents | Where-Object eventData -eq "Initializing Windows Update Agent")[-1].dateString)
                $LastWSUSUpdateUrlFound = ($WUEvents | Where-Object eventData -like "*WSUS server:*$WUServer*")[-1]
                $LastSearchResult = ($WUEvents | Where-Object eventData -like "*END*Finding updates CallerId = Windows Defender*")[-1]
                $LastSearchResult.eventData -match "Exit code = (?<ExitCode>0x\w{8})" | Out-Null
                $LastSearchExitCode = $Matches.ExitCode 

                # Check search result
                if ($LastSearchExitCode -eq "0x00000000"){
                    Write-Host "[+] Update search is successful."
                    $LastSearchUpdateFound = $WUEvents | Where-Object { [datetime]($_.dateString) -gt $LastAgentInitiateTime } | Where-Object eventData -like "*Search ClientId = Windows Defender, Updates found*"
                    $LastSearchUpdateFound.eventData -match "Updates found = (?<UpdatesCount>\d+)" | Out-Null
                    $UpdatesCount = $Matches.UpdatesCount
                }
                else {
                    List-Warning "Last search failed with error: $LastSearchExitCode" 
                    $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.WUServer" -Value "[Error] search failed with error: $LastSearchExitCode"
                    continue
                }

                # Find updates
                if ($UpdatesCount -eq "0"){
                    List-Warning "[!] 0 available update is found from WSUS: $WUServer, it will return 'No Updates Needed'" 
                    $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.WUServer" -Value "[Warning] $UpdatesCount update(s) found"
                    continue
                }
                else {
                    Write-Host "[+] $UpdatesCount update(s) found. Checking if download succeeded."
                    $DownloadEndEvent = $WUEvents | Where-Object eventPid -eq $LastWSUSUpdateUrlFound.eventPid | Where-Object eventType -eq "DownloadManager" | Where-Object eventData -Like "*END*Downloading Updates*Windows Defender*"
                    $DownloadEndEvent.eventData -match "hr = (?<ExitCode>0x\w{8})" | Out-Null
                    $DownloadExitCode = $Matches.Exitcode
                }

                # Check the signature version in the search result
                $SearchedSignatureVersions = $WUEvents | Where-Object eventPid -eq $DownloadEndEvent.eventPid | Where-Object eventData -like "*KB2267602*"
                $SearchedSignatureVersions[-1].eventData -match "KB2267602 \(Version (?<SignatureVersion>\d+\.\d+\.\d+\.\d+)\)" | Out-Null
                $LatestAvailableSignatureVersion = $Matches.SignatureVersion
                
                
                # Check downloads
                if ($LatestAvailableSignatureVersion) {
                    Write-Host "[+] Latest available signature version: $LatestAvailableSignatureVersion"
                } else {
                    Write-Host "[!] Unable to extract the latest available signature version."
                }

                # Check version match
                if ($LatestAvailableSignatureVersion -ne $SignatureVersion){
                    List-Warning "The latest available signature version is $LatestAvailableSignatureVersion, which is different from the current signature version $SignatureVersion"
                }

                $IsDownloadSuccess = $false
                if ($DownloadExitCode -eq '0x00000000'){
                    List-Success "All $UpdatesCount updates are downloaded successfully." 
                    $IsDownloadSuccess = $true
                }
                else {
                    List-Warning "Last download failed with code: $DownloadExitCode" 
                    return
                }

                if ($IsDownloadSuccess){
                    # Install updates
                    $InstallCompleteEvent = ($WUEvents | Where-Object eventPid -eq $LastSearchUpdateFound.eventPid | Where-Object eventType -eq "ComApi" | Where-Object eventData -Like "*Install call complete*")[-1]
                    $InstallationIsCompleted = $InstallCompleteEvent.eventData -match "succeeded = (?<Succeeded>\d+), succeeded with errors = (?<SucceededWithErrors>\d+), failed = (?<Failed>\d+)"
                    Write-Host "Installation completed: $InstallationIsCompleted"
                    if (!$InstallationIsCompleted){
                        List-Warning "Installation completed event is not found."
                        $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.WUServer" -Value "[Warning] Installation not found"
                        continue
                    }

                    if ($Matches.Succeeded -eq $UpdatesCount){
                        Write-Host -ForegroundColor Green "[+] All $UpdatesCount updates are installed successfully." 
                        $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.WUServer" -Value "[Success] Updates success"
                    }
                    elseif($Matches.Succeeded -ne $UpdatesCount){
                        List-Warning "Not all updates are installed successfully." 
                        Write-Output $InstallCompleteEvent.eventData | Out-File $Result -Append
                        $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.WUServer" -Value "[Warning] Updates partially installed."
                    }
                }

        }
        
        if(!$ForceWSUS){
            List-Success "MU/WU can update from Internet." 
            # Check if update is downloaded from MU/WU
            $DownloadEventsFromInternet = $WUEvents | Where-Object eventType -eq "DownloadManager" | Where-Object eventData -like "*download.windowsupdate.com*"
            
            if ($DownloadEventsFromInternet){
                # Client Can download from Internet
                $LastDownloadJobFromInternet = $WUEvents | Where-Object eventPid -eq $DownloadEventsFromInternet[-1].eventPid | Where-Object eventType -eq "DownloadManager"
                $IsLastDownloadSuccess = $LastDownloadJobFromInternet | Where-Object eventData -like "DO job*completed successfully*"
                if ($IsLastDownloadSuccess){
                    List-Success "Signature Update download job from Internet is successful."
                    $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.MU/WU" -Value "[Success] Download success."
                }
                else{
                    List-Warning "Signature Update download from Internet failed."
                    $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.MU/WU" -Value "[Error] Download failed"
                    # To be continued to catch the error.
                }
            }
            else {
                List-Warning "Cannot find download event from Internet. Try to locate search event."
                $SearchEventsFromInternet = $WUEvents | Where-Object eventType -eq "ComApi" | Where-Object eventData -like "*END*Search ClientId = MoUpdateOrchestrator*Updates found*"
                if ($SearchEventsFromInternet){
                    $SearchEventsFromInternet[-1] -match "Updates found=(?<UpdatesCount>\d+)" | Out-Null
                    $UpdateCount = $Matches.UpdatesCount
                    List-Success "$UpdateCount update(s) has been found, but download was not started somehow."
                    $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.MU/WU" -Value "[Error] Update found but not downloaded"
                    # To be continued.
                }
                else{
                    List-Warning "No update search event is found. Please check Windows Update log manually. Possible cause could be bad connectivity to *.update.microsoft.com."
                    $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MicrosoftUpdateServer.MU/WU" -Value "[Error] Update not found."
                }
            }

        }
    }
        'InternalDefinitionUpdateServer'{
            # WSUS
            List-StartSection "[*] Analyzing WSUS source."
            # Get WSUS config
            if (!$WUServer){
                List-Warning "WUServer config is not found. WSUS cannot update."
                continue
            }
            
            # Check WSUS 
            $MpCmdRunSigUpdateEvents = Get-Content "$MpLogFolder\MpCmdRun-NetworkService.json" | ConvertFrom-Json | Where-Object Command -Like "*signature*"
            $MpCmdRunSigUpdateEvents += Get-Content "$MpLogFolder\MpCmdRun-System.json" | ConvertFrom-Json | Where-Object Command -Like "*signature*"
            $WSUSEvents = $MpCmdRunSigUpdateEvents | Where-Object Details -like "*WSUS Update*"
            $LastWSUSEventDetails = $WSUSEvents[-1].Details
            if ($LastWSUSEventDetails -match "Update completed successfully. no updates needed"){
                List-Success "Update completed via WSUS, but no updates needed."
                List-Warning "If signature is not updated, please confirm if WSUS approves the updates."
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "InternalDefinitionUpdateServer" -Value "[Warning] No updates needed"
                # Future plan: check MpSigStub.log or not?
            }
            elseif ($LastWSUSEventDetails -match "Update failed with hr: (?<ErrorCode>0x\w{8})"){
                $LastWSUSUpdateErrorLine = ($LastWSUSEventDetails | Where-Object { $_ -like "*Update failed*"})[-1]
                $LastWSUSUpdateErrorLine -match "Update failed with hr: (?<ErrorCode>0x\w{8})" | Out-Null
                $LastWSUSUpdateErrorCode = $Matches.ErrorCode
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "InternalDefinitionUpdateServer" -Value "Failed with error: $LastWSUSUpdateErrorCode"
                List-Warning "Last WSUS update failed from MpCmdRun log."
                List-Warning $Matches[0]
            }
        }
        'MMPC'{
            # ADL/HTTP
            List-StartSection "Analyzing MMPC source." 
            Write-Host "[*] Analyzing from Event Log ID 2001." 
            $MMPCUpdateFailureEvents = Import-Csv "$MpLogFolder\MpOperationalEvents.csv" | Where-Object EventId -eq "2001" | Where-Object EventDetails -match "Microsoft Malware Protection Center"
            if ($MMPCUpdateFailureEvents) { 
                $LastMMPCUpdateFailureDetails = ($MMPCUpdateFailureEvents[0]).EventDetails
            }
            $LastMMPCUpdateFailureDetails -match "Error code: (?<ErrorCode>0x\w{8})" | Out-Null
            $LastMMPCErrorCode = $Matches.ErrorCode 
            if ($LastMMPCErrorCode){
                List-Warning "Last MMPC update exit with code: $lastMMPCErrorCode"
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MMPC" -Value "[Error] Update error code: $lastMMPCErrorCode"
            }
            $LastMMPCUpdateFailureDetails -match "Error description: (?<ErrorDescription>.*)" | Out-Null
            $LastMMPCErrorDescription = $Matches.ErrorDescription
            if ($LastMMPCErrorDescription){
                List-Warning "Last MMPC Error description: $LastMMPCErrorDescription" 
            }

            # Confirm from MpCmdRun-system.json
            Write-Output "Analyzing from MpCmdRun-system." 
            $MpCmdRunSystemMmpc = Get-Content "$MpLogFolder\MpCmdRun-system.json" | ConvertFrom-Json | Where-Object Command -like "*signature*" | Where-Object Details -match "Direct HTTP"
            if ($MpCmdRunSystemMmpc){
                $LastMmpcUpdateDetails = $MpCmdRunSystemMmpc[-1].Details 
            }
            else{
                List-Warning "[!] No MMPC update is found in MpCmdRun-system" 
            }

            if ($LastMmpcUpdateDetails -match "Update completed succes*"){
                List-Success "Last MMPC update was successful from MpCmdRun-system."
                $MmpcUpdateTime = $MpCmdRunSystemMmpc[-1].StartTime
                List-Success "Last MMPC updated in MpCmdRun-system was: $MmpcUpdateTime"
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MMPC" -Value "[Success] Last MMPC update succeeded"
            }
            elseif ($LastMmpcUpdateDetails -match "Update failed with hr: (?<ErrorCode>0x\w{8})"){
                $LastUpdateErrorLine = ($LastMmpcUpdateDetails | Where-Object { $_ -like "*Update failed*"})[-1]
                $LastUpdateErrorLine -match "Update failed with hr: (?<ErrorCode>0x\w{8})" | Out-Null
                List-Warning "Last MMPC update failed from MpCmdRun-system."
                List-Warning $Matches[0]
                $LastUpdateErrorCode = $Matches.ErrorCode
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "MMPC" -Value "[Error] Last MMPC update failed: $LastUpdateErrorCode"
            }
        }
        'FileShares'{
            # UNC path
            List-StartSection "[*] Analyzing UNC / File shares source" 
            if (!($SignatureConfigs.DefinitionUpdateFileSharesSources)){
                List-Warning "File share source: DefinitionUpdateFileSharesSources is not configured. Cannot update from FileShare. Exit."
                continue
            }
            else{
                $DefinitionUpdateFileSharesSources = $SignatureConfigs.DefinitionUpdateFileSharesSources
            }

            if (!($DefinitionUpdateFileSharesSources -match "\\\\.+")){
                List-Warning "FileShare: $DefinitionUpdateFileSharesSources seems to be a wrong format."
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "FileShares" -Value "[Error] FileShare: $DefinitionUpdateFileSharesSources in wrong format"
                continue
            }
            else{
                List-Info "FileShare location:" $DefinitionUpdateFileSharesSources
            }

            $UNCUpdateEvents = Get-Content "$MpLogFolder\MpCmdRun-System.json" | ConvertFrom-Json | Where-Object Command -Like "*signature*" | Where-Object Details -match "UNC share"
            $LastUNCUpdateEvent = $UNCUpdateEvents[-1]
            if ($LastUNCUpdateEvent.CommandReturn){
                $UNCCommandReturn = LastUNCUpdateEvent.CommandReturn
                List-Warning "Last UNC update failed with error: $UNCCommandReturn"
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "FileShares" -Value "[Error] Failed with error: $UNCCommandReturn"
            }
            else{
                $LastUNCUpdateDetails = $UNCUpdateEvents[-1].Details
            }

            if ($LastUNCUpdateDetails -match "no updates needed"){
                List-Success "Last UNC update returned 'No updates needed.'"
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "FileShares" -Value "[Warning] No updates needed"
            }
            elseif ($LastUNCUpdateDetails -like "*completed succes*"){
                List-Success "Last UNC update completed with success."
                $FallbackOrderResults | Add-Member -MemberType NoteProperty -Name "FileShares" -Value "[Success] Update success"
            }


        }
        Default {
            Write-Host -ForegroundColor Red "[-] Update Source '$UpdateSource' is not valid or in the wrong format. Skip analysis." | Out-File $Result -Append
        }
    }
    $sourceOrder++
    
}
List-EndSection "Log analysis completed."
$FallbackOrderResults | Format-List
$FallbackOrderResults | Format-List | Out-File $Result -Append
List-EndSection "Analyais log is: $Result"
# Final result: list an output of the failure/success reason of each update source.