function Convert-MpOperationalEventLogTxt(
    [string]$path = "$PWD\MpOperationalEvents.txt",
    [string]$outFile = "$PWD\MpOperationalEvents.csv",
    [switch]$expandEventDetails
) {

    $asCsv = $outFile -imatch "\.csv$"
    $asJson = $outFile -imatch "\.json$"

    if (!$asCsv -and !$asJson) {
        Write-Host "Output file type must be either .csv or .json"
        return
    }

    if ((Test-Path $path) -eq $false ) {
        Write-Host "$path is not found! Exit"
        return
    }

    #Read Data
    $eventSeparator = "*****"
    $logData = Get-Content -path $path

    # Some variables
    $result = [collections.arraylist]::new()
    $totalLines = $logData.Count
    $header = @{}
    $eventMessage = [collections.arrayList]::new()
    $eventStartIndex = 0

    for ($i = 0; $i -lt $totalLines) {
        $line = $logData[$i]
        $i++

        if ($line.Length -lt 1) {
            Write-Verbose "empty line. line: $($i - 1)"
            continue
        }

        Write-Progress -Activity "Parsing" -Status "$i of $totalLines lines parsed" -PercentComplete (($i / $totalLines) * 100)

        if ($line.StartsWith($eventSeparator) -or ($i -eq $totalLines)) {
            [void]$result.Add((
                    Format-EventMessage -eventMessage $eventMessage `
                        -fieldList $header `
                        -eventIndex $eventStartIndex `
                        -eventDetails:$expandEventDetails `
                        -asJson:$asJson
                ))

            $eventStartIndex = $i
            [void]$eventMessage.Clear()
            continue
        }

        [void]$eventMessage.Add($line)
    }

    [void]$result.Insert(0, (Format-Headers -headerDictionary $header -eventDetails:$expandEventDetails))
    $jsonFile = ConvertTo-Json $result

    if ($asJson) {
        $jsonFile | Out-File -Encoding utf8BOM -FilePath $outFile
    }
    else {
        $result | Export-Csv -Encoding utf8BOM -Path $outFile
    }

    Write-host "Processed lines: $totalLines events: $($result.Count - 1) File: $outFile"
}

function Format-EventMessage {
    param(
        [collections.arrayList]$eventMessage,
        [hashtable]$fieldList,
        [int]$eventIndex,
        [switch]$eventDetails,
        [switch]$asJson
    )

    if (!$eventMessage) {
        Write-Warning "empty event message. line: $eventIndex"
    }
    $firstLine = $eventMessage[0]
    $firstLinePattern = "(?<Timestamp>.+?)`t(?<Provider>.+?)`t(?<Level>.+?)`t`t`t(?<Id>.+?)`t(?<Machine>.+)"
    $firstLineMatch = [regex]::Match($firstLine, $firstLinePattern)

    if ($firstLineMatch.Success -eq $false) {
        Write-Warning "bad event first line: $eventIndex : $firstLine"
        $timestamp = "unknown"
        $level = "unknown"
        $id = "unknown"
        $machine = "unknown"
    }
    else {
        $timestamp = $firstLineMatch.Groups["Timestamp"].Value
        $level = $firstLineMatch.Groups["Level"].Value
        $id = $firstLineMatch.Groups["Id"].Value
        $machine = $firstLineMatch.Groups["Machine"].Value
        # convert timestamp from g to datetime suitable for kusto ingestion
        # 6/9/2024 12:53:54 PM
        $timestamp = [dateTimeOffset]::Parse($timestamp).ToString("yyyy-MM-ddTHH:mm:ss.fffzzz")
        #$timestamp = [datetime]::Parse($timestamp).ToUniversalTime().tostring("yyyy-MM-ddTHH:mm:ss.fffZ")
    }

    # remove first line with timestamp, level, id, machine
    [void]$eventMessage.RemoveAt(0)
    $eventDescription = $eventMessage[0].Replace("`r`n", "").Trim()

    # remove event description
    [void]$eventMessage.RemoveAt(0)

    # event details
    $eventDetailsEvents = $eventMessage.ToArray()

    if ($eventDetailsEvents.Count -lt 1) {
        Write-Warning "bad event: line: $eventIndex : $eventMessage"
        return
    }

    $message = [ordered]@{
        EventTimestamp   = $timestamp
        ErrorLevel       = $level
        EventId          = $id
        Machine          = $machine
        EventDescription = $eventDescription
    }

    [void]$eventMessage.RemoveAt(0)

    if ($eventDetails) {
        Format-EventMessageDetails -eventDetailsEvents $eventDetailsEvents `
            -eventIndex $eventIndex `
            -message $message `
            -fieldList $fieldList
    }
    else {
        if ($asJson) {
            $eventDetailsJson = [ordered]@{}
            Format-EventMessageDetails -eventDetailsEvents $eventDetailsEvents `
                -eventIndex $eventIndex `
                -message $eventDetailsJson `
                -fieldList $fieldList
            $message.EventDetails = $eventDetailsJson | ConvertTo-Json
        }
        else {
            $message.EventDetails = ([string]::join("`n", $eventMessage.ToArray()))
        }
    }

    return $message
}

function Format-EventMessageDetails {
    param(
        [string[]]$eventDetailsEvents,
        [int]$eventIndex,
        [Collections.Specialized.OrderedDictionary]$message,
        [hashtable]$fieldList
    )

    for ($i = 0; $i -lt $eventDetailsEvents.Count; $i++) {
        $lineDetails = $eventDetailsEvents[$i].split(":", 2)
        if ($lineDetails.Count -lt 1) {
            Write-Warning "bad details: line: $($eventIndex + $i) : $eventDetailsEvents[$i]"
            continue
        }
        $detailsName = $lineDetails[0].Trim().Replace(" ", "-")
        $detailsValue = ""
        if ($lineDetails.Count -gt 1) {
            $detailsValue = $lineDetails[1].Replace("`r`n", "").Trim()
        }
        if ($message.Contains($detailsName)) {
            Write-Warning "Duplicate key: $detailsName : $detailsValue"
        }
        else {
            [void]$message.Add($detailsName, $detailsValue)
        }

        if ($fieldList.Contains($detailsName) -eq $false) {
            [void]$fieldList.Add($detailsName, $detailsName)
        }
    }
}

function Format-Headers {
    param(
        [hashtable]$headerDictionary,
        [switch]$eventDetails
    )

    $header = [ordered]@{
        EventTimestamp   = "Timestamp"
        ErrorLevel       = "Level"
        EventId          = "Id"
        Machine          = "Machine"
        EventDescription = "Description"
    }

    if (!$eventDetails) {
        [void]$header.Add("EventDetails", "Details")
    }
    else {
        foreach ($kvp in $headerDictionary.GetEnumerator() | Sort-Object Name) {
            if ($header.Contains($kvp.Name)) {
                Write-Warning "Duplicate header: $($kvp.Name) : $($kvp.Value)"
            }
            else {
                [void]$header.Add($kvp.Name, $kvp.Value)
            }
        }
    }
    return $header
}Export-ModuleMember -Function Convert-MpOperationalEventLogTxt


function Convert-MacDlpPolicyBin {

    param (
        [string]$Path="$PWD\dlp_policy.bin",
        [string]$OutFile="$PWD\dlp_policy.json"
    )

    if ((Test-Path $Path) -eq $false ) {
        Write-Host "Policy file not found. Exiting."
        exit
    }
    $byteArray = Get-Content -AsByteStream -Path $Path

    $memoryStream = New-Object System.IO.MemoryStream(,$byteArray)
    $deflateStream = New-Object System.IO.Compression.DeflateStream($memoryStream,  [System.IO.Compression.CompressionMode]::Decompress)
    $streamReader =  New-Object System.IO.StreamReader($deflateStream, [System.Text.Encoding]::utf8)
    $policyStr = $streamReader.ReadToEnd()
    $policy = $policyStr | ConvertFrom-Json


    $policyBodyCmd = ($policy.body | ConvertFrom-Json).cmd 
    if ($policyBodyCmd) {Set-Content -Path $OutFile $policyBodyCmd} 
}
Export-ModuleMember -Function Convert-MacDlpPolicyBin

function Convert-UnixTime([string]$unixTime){
    # Unix time converting
    $base = [datetime]'1970-01-01 00:00:00'
    $timezone = (Get-TimeZone).BaseUtcOffset
    $baseTime = $base + $timezone

    switch ($unixTime.length) {
        10 { $converter = $baseTime.Addseconds($unixTime)}
        13 { $converter = $baseTime.AddMilliseconds($unixTime)}
        16 { $converter = $baseTime.AddMicroseconds($unixTime)}
        Default { Write-Host "Not a valid Unix Time!";break}
    }
    return $converter
}
Export-ModuleMember -Function Convert-UnixTime

function Convert-Wdavhistory {

    param(
        $Path="$PWD\wdavhistory",
        $OutFile="$PWD\wdavhistory.log"
    )

    $Logdata = Get-Content -Raw $Path | ConvertFrom-Json

    $Logdata.scans | ForEach-Object {
        $_.endTime = Convert-UnixTime $_.endTime
        $_.startTime = Convert-UnixTime $_.startTime
        $duration = $_.endTime - $_.startTime
        $_ | Add-Member -NotePropertyName scanDuration -NotePropertyValue $duration.ToString()
        $_.threats = $_.threats | Out-String
    }

    $Logdata.scans | Format-Table -Property startTime, endTime, scanDuration, filesScanned, scheduled, state, type, threats| Out-File $OutFile
}
Export-ModuleMember -Function Convert-Wdavhistory

function Convert-MdavRtpDiag() {
    param(
        [Parameter(Mandatory=$true)]$rtplog,
        $OutFile = "$rtplog.txt"
    )
    if (Test-Path $rtpLog){
        $val = Get-Content $rtpLog | ConvertFrom-Json
        $val.counters | ForEach-Object {$_.totalFilesScanned = [int32]$_.totalFilesScanned} # Convert totalFileScanned type from string to int.
        $val.counters | Where-Object {$PSItem.totalFilesScanned -gt 0} | Select-Object id, name, totalFilesScanned, path | Sort-Object -Property totalFilesScanned -Descending | Format-Table | Out-File -Force $OutFile
    }
    else {
        Write-Host "Cannot find the file $rtplog. Exit."
        return
    }
    
    Write-Host "File saved as $OutFile successfully."
}
Export-ModuleMember -Function Convert-MdavRtpDiag

function Convert-MpRegistrytxtToJson {
    param(
        [string]$Path="$PWD\MpRegistry.txt",
        [string]$OutFile="$PWD\MpRegistry.json"
    )

    $lines = Get-Content $Path
    $json_result = [ordered]@{}

    $this_line = 0
    $depth = 1
    $regs = ""
    $key = ""
    $value = ""
    
    :DefenderAV foreach ($line in $lines) {
        # Match the root category
        if ($line.StartsWith("Current configuration options for location")){
            $policy = ($line -split '"')[1]
            $json_result.$policy = [ordered]@{}
        }

        elseif ($line.StartsWith("Windows Setup keys from")) {
            $this_line++
            break DefenderAV # exit when Defender AV regs are finished
        }
        
        # Match a category like '[*]'
        elseif ($line -match '^\[(.+)\]$') {
            $regs = $matches[1] -split "\\"
            $depth = $regs.Length
            
            # Length of 'keys' is the depth of json
            switch ($depth){
                1 { $json_result.$policy.($regs[0]) = [ordered]@{}; break }
                2 { $json_result.$policy.($regs[0]).($regs[1]) = [ordered]@{}; break }
                {$line -eq '[NIS\Consumers\IPS]'} { $json_result.$policy.($regs[0]).($regs[1]) = [ordered]@{} }
                3 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]) = [ordered]@{}; break }
                4 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).($regs[3]) = [ordered]@{}; break }
            }
        } 

        # Match a line like *key* : *value*
        elseif ($line -match '^\s{4}(\S+)\s+\[.+\]\s+:\s(.+)$') {
            $key = $matches[1]
            $value = $matches[2]
            if ($value -in ("<NO VALUE>","(null)")) { $value = "" }
            switch ($depth){
                1 { $json_result.$policy.($regs[0]).$key = $value; break }
                2 { $json_result.$policy.($regs[0]).($regs[1]).$key = $value; break}
                3 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).$key = $value; break }
                4 { $json_result.$policy.($regs[0]).($regs[1]).($regs[2]).($regs[3]).$key = $value; break }
            }
        }

        # Try to merge Device Control policy into single line. It will break when the DC policy format is not expected.
        elseif ($line -match '\s{2}<.+>$' ) {
            $json_result.$policy.($regs[0]).($regs[1]).$key += $line
        }

        $this_line++
        }

        # Deal with other registry keys
        # Initializing
        $json_result["Others"]=[ordered]@{}
        $json_result["WindowsUpdate"]=[ordered]@{}
        $policy = "Others"
        $reg_path = ""


        for ($i = $this_line; $i -lt $lines.Count; $i++) {

            switch -Wildcard ($lines[$i]){
                '`[HKEY_*WindowsUpdate*`]' { 
                    $policy = "WindowsUpdate"
                    $reg_path = $lines[$i].Substring(1,$lines[$i].Length-2)
                    $json_result.$policy.$reg_path = [ordered]@{}
                    break
                }
                '`[HKEY*`]' { 
                    $policy = "Others"
                    $reg_path = $lines[$i].Substring(1,$lines[$i].Length-2)
                    $json_result.$policy.$reg_path = [ordered]@{}
                    break
                }
                '*`[REG_*`]*' {
                    if ($lines[$i] -match '^\s{4}(\S+)\s+\[.+\]\s+:\s(.+)$'){
                        $key = $matches[1]
                        $value = $matches[2]
                        $json_result.$policy.$reg_path.$key = $value
                    }

                }
            }
        }

    $json = $json_result | ConvertTo-Json -Depth 4 -WarningAction Ignore | ForEach-Object {
        [Regex]::Replace($_, "\\u(?<Value>[a-fA-F0-9]{4})", { param($m) ([char]([int]::Parse($m.Groups['Value'].Value, [System.Globalization.NumberStyles]::HexNumber))).ToString() } )}       
        # Converto-Json in PowerShell 5.1 does not have escape handling options, so we do it manually.
        # Reference: https://stackoverflow.com/questions/47779157/convertto-json-and-convertfrom-json-with-special-characters/47779605#47779605
    $json | Out-File -FilePath $OutFile -Force
}
Export-ModuleMember -Function Convert-MpRegistrytxtToJson

function Update-DefenderToolbox {
    # Global variables
    $PSVersion = [string]$PSVersionTable.PSVersion
    $PSModuleUserProfile = "\WindowsPowerShell\Modules"
    if ($PSVersion -gt "7.1"){$PSModuleUserProfile = "\PowerShell\Modules"}
    $ModuleFolder = [System.Environment]::GetFolderPath('MyDocuments') + $PSModuleUserProfile
    $ModuleName = "Defender-Toolbox"
    $ModuleFile = "$ModuleName.psm1"
    $ModuleManifestFile = "$ModuleName.psd1"
    $result = $false

    function Get-LatestVersion{
        # $versionListFileUri = "http://127.0.0.1/version_list" # for local test (python3 -m http.server 8080)
        $versionListFileUri = "https://raw.githubusercontent.com/typicaldao/$ModuleName/main/version_list"
        try {
            $version = Invoke-RestMethod -Uri $versionListFileUri
            $latest_version = $version.Split()[-1]
            return [string]$latest_version
        }
        catch {
            Write-Host "Failed to get remote version."
            Write-Host -ForegroundColor Red "Error: $_"
            return "0"
        }
    }

    function Get-LocalVersion{
        if (Test-Path $ModuleFolder\$ModuleName){
            try {
                Import-Module $ModuleName -ErrorAction Stop -DisableNameChecking # Try to import the module
                $local_version = (Get-Module -Name $ModuleName).Version
                # Remove-Module -Name $ModuleName # The module should not be removed when used as a local module.
                return $local_version # Returns System.Version. Use ToString() to convert type.
            }
            catch {
                Write-Host "Failed to import module $ModuleName."
            }
        }
        else {
            return "0.0" # Returns a 0.0 version if module has never been installed.
        }
    }

    function Download-DefenderToolbox([string]$version) {
        # Module links
        $ModuleFileUri = "https://raw.githubusercontent.com/typicaldao/$ModuleName/main/Modules/$version/$ModuleFile"
        $ModuleManifestUri = "https://raw.githubusercontent.com/typicaldao/$ModuleName/main/Modules/$version/$ModuleManifestFile"
        
        Write-Host "Try to download the latest version from GitHub to your temp folder."

        try {
            Invoke-RestMethod -Uri $ModuleFileUri -OutFile $env:TEMP\$ModuleFile
            Invoke-RestMethod -Uri $ModuleManifestUri -OutFile $env:TEMP\$ModuleManifestFile
        }
        catch {
            Write-Host -ForegroundColor Red "Downloading failed:"
            Write-Host $_
            return $false
        }

        Write-Host -ForegroundColor Green "Download $ModuleFile and $ModuleManifestfile at $env:TEMP successfully."
        return $true
    }

    function Copy-ModuleFiles([string]$version){
        # Do this when the download is successful.
        # OneDrive folder needs to be confirmed in the future. To be continued.

        # Create module folder with versions if it does not exists.
        if(-Not (Test-Path $ModuleFolder\$ModuleName\$version)){
            New-Item -ItemType Directory -Path $ModuleFolder\$ModuleName -Name $version
        }
        
        # Confirm module installation path.
        if ($ModuleFolder -in $env:PSModulePath.Split(";")){
            Write-Host "Trying to install $ModuleName version $version at path: $ModuleFolder."
            try {
                Copy-Item -Path $env:TEMP\$ModuleFile -Destination $ModuleFolder\$ModuleName\$version
                Copy-Item -Path $env:TEMP\$ModuleManifestFile -Destination $ModuleFolder\\$ModuleName\$version
            }
            catch {
                Write-Host -ForegroundColor Red "Failed to copy files."
                Remove-Item $ModuleFolder\$ModuleName\$version # clean up the folder.
                return $false
            }
        }
        else{
            Write-Host -ForegroundColor Red "Module folder $ModuleFolder is not inside PowerShell module path."
            Remove-Item $ModuleFolder\$ModuleName\$version # clean up the folder.
            Write-Host -ForegroundColor Yellow "Please manually copy the files $env:TEMP\$ModuleFile and $env:TEMP\$ModuleManifestFile to one of the folders below."
            Write-Host $env:PSModulePath.Split(";")
            return $false
        }  
        
        Write-Host -ForegroundColor Green "Module files of $ModuleName version $version are copied to the module folder."
        return $true
    }

    function Update-PsUserProfile {
        $command = "`nImport-Module -Name $ModuleName -DisableNameChecking"
        $imported = $false
        if (Test-Path $PROFILE){
            $ProfileContent = Get-Content $PROFILE
            $ProfileContent | ForEach-Object {
                if ($_.Trim() -eq $command.Trim()) { 
                    $imported = $true
                    Write-Host "$ModuleName has been imported automatically."
                    return 
                }
            }
            if (!$imported){
                Write-Host -ForegroundColor Cyan "Trying to add a new line to import module in your existing profile. Please confirm."
                $command | Out-File $PROFILE -Append -Confirm
            }
        }
        else{
            Write-Host -ForegroundColor Cyan "You do not have a PowerShell profile. Would you create one and import the module automatially?"
            New-Item -ItemType File -Value $command -Confirm $PROFILE 
        }
    }

    # Main
    $local_version = Get-LocalVersion
    $version = Get-LatestVersion

    if ($version -gt $local_version){
        $result = Download-DefenderToolbox($version)
        $result = Copy-ModuleFiles($version)
    }
    elseif ($version -eq $local_version) {
        Write-Host "You have already installed the latest version: $version"
        $result = $true
    }

    if ($result) { Update-PsUserProfile }   # Will add 'Import-Module -Name Defender-Toolbox -DisableNameChecking' in Profile so the module is automatically loaded.
}
Export-ModuleMember -Function Update-DefenderToolbox

function Read-WindowsUpdateLog{
    param(
    [string]$Path="$PWD\WindowsUpdate.log",
    [string]$OutFile="$PWD\WindowsUpdateLog.csv"
)

$results = [array][PSCustomObject]@()
$lines = Get-Content $Path

$TotalLines = $lines.Length
$i = 0

foreach ($line in $lines){
    if ($line.Length -gt 20){
        $results += [PSCustomObject]@{
            dateString = $line.Substring(0,27)
            eventPid = $line.Substring(28,5).Trim()
            eventTid = $line.Substring(34,5).Trim()
            eventType = $line.Substring(40,15).Trim()
            eventData = $line.Substring(56)
        }
        $i++
        Write-Progress -Activity "Parsing" -Status "$i of $totalLines lines parsed" -PercentComplete (($i / $TotalLines) * 100)
    }
}

    $results | Export-Csv $OutFile
}
Export-ModuleMember -Function Read-WindowsUpdateLog

function Read-MpCmdRunLog{
    
[cmdletbinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$logFilePath,
  [string]$OutFile,
  [switch]$errorsOnly,
  [switch]$quiet
)

$mpCmdRunLogResults = [ordered]@{}
$recordSeparator = '-------------------------------------------------------------------------------------'
$global:records = [collections.arrayList]::new()
$scriptName = "$psscriptroot\$($MyInvocation.MyCommand.Name)"
$recordIdentifier = 'MpCmdRun: '
$startOfRecord = "$($recordIdentifier)Command Line:"
$endOfRecord = "$($recordIdentifier)End Time:"
$mpCmdRunExe = 'MpCmdRun.exe'

function main() {
  try {
    if (!(Test-Path $logFilePath)) {
      Get-Help $scriptName -Examples
      Write-Host "File not found: $logFilePath"
      return $null
    }

    $mpCmdRunLogResults = Read-Records $logFilePath

    $eventTypes = $mpCmdRunLogResults.Command | Group-Object | Sort-Object | Select-Object Count, Name
    Write-Host "Event Types: $($eventTypes | Out-String)" -ForegroundColor Green

    $errorEventTypes = $mpCmdRunLogResults.Where({ $psitem.Level -ieq 'warning' }).Command | Group-Object | Sort-Object | Select-Object Count, Name
    
    if ($errorEventTypes) {
      Write-Host "Event Types with Warnings: $($errorEventTypes | out-string)" -ForegroundColor Yellow
    }
    
    $errorEventTypes = $mpCmdRunLogResults.Where({ $psitem.Level -ieq 'error' }).Command | Group-Object | Sort-Object | Select-Object Count, Name
    if ($errorEventTypes) {
      Write-Host "Event Types with Errors: $($errorEventTypes | out-string)" -ForegroundColor Red
    }

    $global:records | ConvertTo-Json | Out-File $OutFile
  }
  catch {
    write-verbose "variables:$((get-variable -scope local).value | convertto-json -WarningAction SilentlyContinue -depth 2)"
    write-host "exception::$($psitem.Exception.Message)`r`n$($psitem.scriptStackTrace)" -ForegroundColor Red
    return $null
  }
}

function Find-RecordMatches([collections.arrayList]$record, [string]$pattern, [bool]$ignoreCase = $true, [Text.RegularExpressions.RegexOptions]$regexOptions = [Text.RegularExpressions.RegexOptions]::None) {
  Write-Verbose "Find-RecordMatches([string]$record,[string]$pattern, [bool]$ignoreCase, [Text.RegularExpressions.RegexOptions]$regexOptions)"
  $line = [string]::Join("`n", $record.ToArray())
  return Find-Matches -line $line -pattern $pattern -ignoreCase $ignoreCase -regexOptions $regexOptions
}

function Find-Matches([string]$line, [string]$pattern, [bool]$ignoreCase = $true, [Text.RegularExpressions.RegexOptions]$regexOptions = [Text.RegularExpressions.RegexOptions]::None) {
  Write-Verbose "Find-Matches([string]$line, [string]$pattern, [bool]$ignoreCase, [Text.RegularExpressions.RegexOptions]$regexOptions)"
  if (!$line) {
    Write-Error "Regex-Match: No line provided"
    throw
    return $null
  }

  if (!$pattern) {
    Write-Error "Regex-Match: No pattern provided"
    throw
    return $null
  }

  if ($ignoreCase) {
    $regexOptions = $regexOptions -bor [Text.RegularExpressions.RegexOptions]::IgnoreCase
  }

  Write-Verbose "Regex-Match: [regex]::Match($line, $pattern, $regexOptions)"
  $regexMatchCollection = [regex]::Matches($line, $pattern, $regexOptions)

  if ($regexMatchCollection.Count -eq 0) {
    Write-Verbose "Regex-Match: No match found for pattern: $pattern"
    return $null
  }
  elseif ($regexMatchCollection.Count -gt 1) {
    Write-Verbose "Regex-Match: Multiple matches found for pattern: $pattern returning collection"
    return $regexMatchCollection
  }

  foreach ($regexMatch in $regexMatchCollection) {
    if ($regexMatch.Success) {
      if ($regexMatch.Groups.Count -eq 1 -and [string]::IsNullOrEmpty($regexMatch.Groups[0].Value)) {
        Write-Verbose "Regex-Match: No groups found. line: $line pattern: $pattern"
        return $null
      }

      # create key value pair
      $groups = @{ }
      foreach ($group in $regexMatch.Groups) {
        Write-Verbose "Regex-Match: Group: $($group.Name) = $($group.Value)"
        [void]$groups.Add($group.Name.Trim(), $group.Value.Trim())
      }

      return $groups
    }
  }

  Write-Verbose "Regex-Match: No match found for pattern: $pattern"
  return $null
}

function Format-Timestamp([string]$timestamp, [string]$timeFormat = "yyyy-MM-ddTHH:mm:ss.fffzzz") {
  return [dateTimeOffset]::Parse($timestamp).ToString($timeFormat)
}

function Get-Commands([string]$commandLine, [collections.specialized.orderedDictionary]$record) {
  Write-Verbose "Get-Commands([string]$commandLine)"
  $commandOptions = [collections.Arraylist]::new()
  $commandString = (Find-Matches -line $commandLine -pattern "$mpCmdRunExe`"?(?<commandString>.+)")

  if (!$commandString) {
    return $commandOptions
  }

  $record.CommandLine = $commandLine.Replace('Command Line: ', '').Trim()
  $commands = $commandString['commandString'].Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
  $optionName = ''

  foreach ($command in $commands) {
    if ($command -eq $commands[0]) {
      $record.Command = $command.Replace('-', '').Trim()
      continue
    }
    if ($command.StartsWith('-')) {
      $optionName = $command.Replace('-', '').Trim()
      $record.CommandOptions += @{$optionName = "" }
    }
    else {
      $record.CommandOptions[$optionName] += $command.Trim()
    }
  }
}

function New-Record() {
  $record = [ordered]@{
    'StartTime'      = ''
    'EndTime'        = ''
    'Level'          = 'Information' # Information, Warning, Error
    'CommandLine'    = ''
    'Command'        = ''
    'CommandOptions' = [ordered]@{}
    'CommandReturn'  = ''
    'Details'        = @()
  }
  return $record
}

function Read-PropertyNameValue(
  [collections.arrayList]$record,
  [string]$propertyName,
  [string]$propertyNamePrefix = '.+?',
  [string]$separator = ':',
  [switch]$all
) {
  Write-Verbose "Read-PropertyNameValue([collections.arrayList]$($record.Count), [string]$propertyName, [string]$propertyNamePrefix, [string]$separator)"
  $kvpCollection = [collections.arrayList]::new()
  $pattern = "$($propertyNamePrefix)(?<propertyName>$($propertyName))\s*?$separator\s*?(?<propertyValue>.*)"

  foreach ($line in $record) {
    $regexMatch = Find-Matches -line $line -pattern $pattern
    if ($regexMatch) {
      $propertyName = $regexMatch['propertyName']
      $propertyValue = $regexMatch['propertyValue'].Trim()
      Write-Verbose "Property name value found: $propertyName = $propertyValue"
      $kvp = @{
        'PropertyName'  = $propertyName
        'PropertyValue' = $propertyValue
      }
      if ($all) {
        [void]$kvpCollection.Add($kvp)
      }
      else {
        return $kvp
      }
    }
  }

  if ($all) {
    return $kvpCollection
  }
  Write-Warning "Property name value not found: $propertyName"
  return $null
}

function Read-Record([collections.arrayList]$record, [int]$index) {
  Write-Verbose "Read-Record([collections.arrayList]$($record.Count), [int]$index)"
  #$newRecord = Read-RecordMetaData -record $record -newRecord (New-Record)
  $newRecord = New-Record
  $returnCode = Find-RecordMatches -record $record -pattern "MpCmdRun.exe: hr = (?<returnCode>.+)"
  if ($returnCode) {
    $newRecord.CommandReturn = $returnCode['returnCode']
  }

  if (Find-RecordMatches -record $record -pattern "Error|Failed|0x[0-9A-Fa-f]8") {
    $newRecord.Level = 'Error'
  }
  elseif (Find-RecordMatches -record $record -pattern "Warning") {
    $newRecord.Level = 'Warning'
  }
  else {
    $newRecord.Level = 'Information'
  }

  for ($recordIndex = 0; $recordIndex -lt $record.Count; $recordIndex++) {
    $line = $record[$recordIndex].Trim()
    if (!$line) { continue }

    if ($recordIndex -eq 0) {
      # Read first line for Command Line
      $line = $line.Replace($recordIdentifier, '')
      Get-Commands -commandLine $line -record $newRecord
      continue
    }
    elseif ($recordIndex -eq 1) {
      # Read second line for Start Time
      $line = $line.Replace('Start Time: ', '').Trim()
      $newRecord.StartTime = Format-Timestamp -timestamp $line
      continue
    }
    elseif ($recordIndex -eq $record.Count - 1) {
      # Read last line for End record
      # unlike Start record, End record is not always at the start of the last line
      $endTime = Find-Matches -line $line -pattern "(?<previousLine>.*?)$($recordIdentifier)End Time: (?<timestamp>.+)"
      if ($endTime -and $endTime['timestamp']) {
        $newRecord.EndTime = Format-Timestamp -timestamp $endTime['timestamp']
      }
      
      if ($endTime -and $endTime['previousLine']) {
        $newRecord.Details += $line
      }

      continue
    }

    $newRecord.Details += $line
  }

  if ($errorsOnly -and $newRecord.Level -ine 'Error') {
    return $null
  }

  if ($quiet -eq $false) {
    $consoleColor = [System.ConsoleColor]::White
    switch ($newRecord.Level) {
      'Error' { $consoleColor = [System.ConsoleColor]::Red }
      'Warning' { $consoleColor = [System.ConsoleColor]::Yellow }
      'Information' { $consoleColor = [System.ConsoleColor]::Green }
    }
    Write-Host "returning record $($records.Count): $($newRecord | out-string)" -ForegroundColor $consoleColor
  }

  return $newRecord
}

function Read-Records($logFilePath) {
  $streamReader = [System.IO.StreamReader]::new($logFilePath, [System.Text.Encoding]::Unicode)
  $inRecord = $false
  $index = 0
  $record = [collections.arrayList]::new()

  while ($streamReader.EndOfStream -eq $false) {
    $line = $streamReader.ReadLine().Trim()
    Write-Verbose "Read-Records: $line"
    $index++

    if ($line.Length -eq 0) {
      continue
    }

    # remove unknown unicode characters outside of ASCII range
    $line = [string]::Join('', ($line.ToCharArray() | Where-Object { [int]$psitem -ge 32 -and [int]$psitem -le 126 }))

    #if ($line.StartsWith($startOfRecord) -and !$inRecord) {
    if ((Find-Matches -line $line -pattern $startOfRecord) -and !$inRecord) {
      # start new record
      $inRecord = $true

      [void]$record.Add($line)
    }
    elseif ((Find-Matches -line $line -pattern $endOfRecord) -and $inRecord) {
      #elseif ($line.StartsWith($endOfRecord) -and $inRecord) {
      $inRecord = $false
      # add record to results
      [void]$record.Add($line)
      [void]$records.Add((Read-Record $record $index))
      [void]$record.Clear()
      continue
    }
    elseif ($line -eq $recordSeparator) {
      continue
    }
    elseif ($inRecord) {
      [void]$record.Add($line)
    }
    else {
      Write-Warning "Unknown Record log file format index: $index line: $line"
    }
  }
  return $records
}

main
} Export-ModuleMember -Function Read-MpCmdRunLog

function Get-SignatureUpdateAnalysis{
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
# Time format could be different. Try to display the value directly.

try {
    $AVSignatureApplied = Convert-RegDateTime $SignatureConfigs."AVSignatureApplied" # Represents the datetime stamp when AV Signature version was created
    $SignaturesLastUpdated = Convert-RegDateTime $SignatureConfigs."SignaturesLastUpdated"
    $LastFallbackTime = Convert-RegDateTime $SignatureConfigs."LastFallbackTime"
}
catch {
    $AVSignatureApplied = $SignatureConfigs."AVSignatureApplied"
    $SignaturesLastUpdated = $SignatureConfigs."SignaturesLastUpdated"
    $LastFallbackTime = $SignatureConfigs."LastFallbackTime"
}

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
                List-Warning "No MMPC update is found in MpCmdRun-system" 
            }

            if ($LastMmpcUpdateDetails -match "Update completed succes*"){
                List-Success "Last MMPC update was successful from MpCmdRun-system."
                $MmpcUpdateTime = $MpCmdRunSystemMmpc[-1].StartTime
                List-Success "Last MMPC updated in MpCmdRun-system was: $MmpcUpdateTime"
            }
            elseif ($LastMmpcUpdateDetails -match "Update failed with hr: (?<ErrorCode>0x\w{8})"){
                $LastUpdateErrorLine = ($LastMmpcUpdateDetails | Where-Object { $_ -like "*Update failed*"})[-1]
                $LastUpdateErrorLine -match "Update failed with hr: (?<ErrorCode>0x\w{8})" | Out-Null
                List-Warning "Last MMPC update failed from MpCmdRun-system."
                List-Warning $Matches[0]
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
} Export-ModuleMember -Function Get-SignatureUpdateAnalysis