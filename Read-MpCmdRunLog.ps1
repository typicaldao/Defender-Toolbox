<#
.SYNOPSIS
    Reads the MpCmdRun*.log files and returns the results as an array of PowerShell objects.

.DESCRIPTION
    Reads the MpCmdRun*.log files and returns the results as an array of PowerShell objects.
    Results are also stored in the global variable $global:mpCmdRunLogResults.
    To troubleshoot, use -Verbose to see additional information.

.NOTES
    File Name      : Read-MpCmdRunLog.ps1
    version        : 0.1

.EXAMPLE
    C:\'Program Files'\'Windows Defender'\MpCmdRun.exe -GetFiles
    copy 'C:\ProgramData\Microsoft\Windows Defender\Support\MpSupportFiles.cab'
    md $pwd\MpSupportFiles
    expand -R -I $pwd\MpSupportFiles.cab -F:* $pwd\MpSupportFiles

    To generate the mpCmdRun.log file

.EXAMPLE
    .\Read-mpCmdRunLog.ps1 -logFilePath $pwd\MpSupportFiles\mpCmdRun.log

    Reads the mpCmdRun.log file and returns the results as a PowerShell object.

.PARAMETER logFilePath
    The path to the mpCmdRun.log file.

.PARAMETER errorsOnly
    Show only errors.

.PARAMETER quiet
    Do not show any output.

.PARAMETER verbose
    Show additional information for troubleshooting.
#>
[cmdletbinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$logFilePath,
  [switch]$errorsOnly,
  [switch]$quiet
)

$global:mpCmdRunLogResults = [ordered]@{}
$recordSeparator = '-------------------------------------------------------------------------------------'
$global:records = [Collections.ArrayList]::New()
$scriptName = "$psscriptroot\$($MyInvocation.MyCommand.Name)"
$recordIdentifier = 'MpCmdRun: '
$startOfRecord = "$($recordIdentifier)Command Line:"
$endOfRecord = "$($recordIdentifier)End Time:"
$mpCmdRunExe = 'MpCmdRun.exe'

function Main() {
  try {
    if (!(Test-Path $logFilePath)) {
      Get-Help $scriptName -Examples
      Write-Error "File not found: $logFilePath"
      return $null
    }

    $global:mpCmdRunLogResults = Read-Records $logFilePath
    $eventTypes = $global:mpCmdRunLogResults.Command | Group-Object | Sort-Object | Select-Object Count, Name
    Write-Host "Event Types: $($eventTypes | Out-String)" -ForegroundColor Green

    $errorEventTypes = $global:mpCmdRunLogResults.Where({ $psitem.Level -ieq 'warning' }).Command | Group-Object | Sort-Object | Select-Object Count, Name
    
    if ($errorEventTypes) {
      Write-Host "Event Types with Warnings: $($errorEventTypes | out-string)" -ForegroundColor Yellow
    }
    
    $errorEventTypes = $global:mpCmdRunLogResults.Where({ $psitem.Level -ieq 'error' }).Command | Group-Object | Sort-Object | Select-Object Count, Name
    if ($errorEventTypes) {
      Write-Host "Event Types with Errors: $($errorEventTypes | out-string)" -ForegroundColor Red
    }

    Write-Host "Results saved to `$global:mpCmdRunLogResults"
    return $global:mpCmdRunLogResults
  }
  catch {
    write-host "exception::$($psitem.Exception.Message)`r`n$($psitem.scriptStackTrace)" -ForegroundColor Red
    return $null
  }
}

function Find-RecordMatches([Collections.ArrayList]$record, [string]$pattern, [bool]$ignoreCase = $true, [Text.RegularExpressions.RegexOptions]$regexOptions = [Text.RegularExpressions.RegexOptions]::None) {
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

      $groups = @{}
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
  $commandOptions = [Collections.ArrayList]::New()
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

function Read-Record([Collections.ArrayList]$record, [int]$index) {
  Write-Verbose "Read-Record([Collections.ArrayList]$($record.Count), [int]$index)"
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
    $consoleColor = [ConsoleColor]::White
    switch ($newRecord.Level) {
      'Error' { $consoleColor = [ConsoleColor]::Red }
      'Warning' { $consoleColor = [ConsoleColor]::Yellow }
      'Information' { $consoleColor = [ConsoleColor]::Green }
    }
    Write-Host "returning record $($records.Count): $($newRecord | out-string)" -ForegroundColor $consoleColor
  }

  return $newRecord
}

function Read-Records($logFilePath) {
  $streamReader = [IO.StreamReader]::New($logFilePath, [Text.Encoding]::Unicode)
  $inRecord = $false
  $index = 0
  $record = [Collections.ArrayList]::New()

  while ($streamReader.EndOfStream -eq $false) {
    $line = $streamReader.ReadLine().Trim()
    Write-Verbose "Read-Records: $line"
    $index++

    if ($line.Length -eq 0) {
      continue
    }

    $line = Remove-UnicodeCharacters -line $line
    if ((Find-Matches -line $line -pattern $startOfRecord) -and !$inRecord) {
      # start new record
      $inRecord = $true

      [void]$record.Add($line)
    }
    elseif ((Find-Matches -line $line -pattern $endOfRecord) -and $inRecord) {
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
  
  $streamReader.Close()
  return $records
}

function Remove-UnicodeCharacters([string]$line) {
  $line = [Text.RegularExpressions.Regex]::Replace($line, "[^\u0000-\u007F]", "")
  return $line
}

Main