<#
.SYNOPSIS
    Merges WindowsUpdate ETL files into a single list of events

.DESCRIPTION
    This script reads WindowsUpdate ETL files.
    The script uses pktmon or netsh to convert the ETL files to CSV format.
    The script then reads the CSV files and parses the entries into a global custom object $global:windowsUpdateEntries
    The custom object contains the following properties:
    - Time: The time the event was created in UTC
    - PID: The process ID in decimal format
    - TID: The thread ID in decimal format
    - Level: The event level (Critical, Error, Warning, Information, Verbose)
    - Keyword: The event keyword
    - Provide: The provider name
    - Event: The event name
    - Info: The event message

.NOTES
File Name      : Read-WindowsUpdate.ps1
version        : 0.2

.EXAMPLE
    C:\'Program Files'\'Windows Defender'\MpCmdRun.exe -GetFiles
    copy 'C:\ProgramData\Microsoft\Windows Defender\Support\MpSupportFiles.cab'
    md $pwd\MpSupportFiles
    expand -R -I $pwd\MpSupportFiles.cab -F:* $pwd\MpSupportFiles

    To generate the MpSupportFiles directory

.EXAMPLE
    Read-WindowsUpdate.ps1 -mpSupportFilesPath $pwd\MpSupportFiles

    To parse the WindowsUpdate ETL files in the MpSupportFiles directory

.PARAMETER mpSupportFilesPath
    The path to the directory containing the WindowsUpdate ETL files

.PARAMETER etlFileFilter
    The filter to apply to the ETL files in the directory
    default: "WindowsUpdate*.etl"

.PARAMETER useNetsh
    Use netsh to convert the ETL files to CSV format
    default: $false

.PARAMETER outputFile
    The path to save the parsed entries to a CSV or JSON file
#>
[cmdletbinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$mpSupportFilesPath, # Path to the directory containing the WindowsUpdate ETL files
  [string]$etlFileFilter = "WindowsUpdate*.etl",
  [switch]$useNetsh, # Using pktmon to parse the ETL files is about 33% faster than netsh but not available on all systems
  [string]$outputFile #= "$pwd\WindowsUpdate.csv" # if specified, the script will save the parsed entries to a CSV file or JSON file
)

$global:windowsUpdateEntries = [System.Collections.ArrayList]::New()
$scriptName = "$psscriptroot\$($MyInvocation.MyCommand.Name)"

class DateTimeComparer : System.Collections.IComparer {
  [int] Compare($x, $y) {
    return [DateTime]::Compare($x.time, $y.time)
  }
}

function Main() {
  try {
    if (!(Test-Path $mpSupportFilesPath)) {
      Get-Help $scriptName -Examples
      Write-Error "The specified path does not exist: $mpSupportFilesPath"
      return $null
    }

    $etlFiles = Get-ChildItem -Path $mpSupportFilesPath -Filter $etlFileFilter -Recurse
    $usepktmon = !$useNetsh -and ($null -ne (Get-Command -Name pktmon -ErrorAction SilentlyContinue))

    if ($etlFiles.Count -gt 0) {
      foreach ($etlFile in $etlFiles) {
        Write-Progress -Activity "Processing ETL files" -Status "Processing $etlFile" -PercentComplete (($etlFiles.IndexOf($etlFile) / $etlFiles.Count) * 100)
        [void]$global:windowsUpdateEntries.AddRange(@(Format-EtlFile -fileName $etlFile.FullName -usepktmon $usepktmon))
      }
    }
    else {
      Write-Error "No ETL files found in the specified directory: $mpSupportFilesPath"
      return $null
    }

    Write-Host "sorting entries" -ForegroundColor Cyan
    $global:windowsUpdateEntries.Sort([DateTimeComparer]::new())

    Save-ToFile -outputFile $outputFile
    $levelGroups = $global:windowsUpdateEntries.Level | Group-Object | Sort-Object | Select-Object Count, Name

    Write-Host "Level Counts:$($levelGroups| out-string)" -ForegroundColor Cyan
    Write-Host "Total entries: $($global:windowsUpdateEntries.Count)"
    Write-Host "Entries saved to `$global:windowsUpdateEntries"
    
    return $global:windowsUpdateEntries
  }
  catch {
    Write-Host "exception::$($psitem.Exception.Message)`r`n$($psitem.scriptStackTrace)" -ForegroundColor Red
    return $null
  }
}

function Convert-ToDecimal([string]$hex) {
  return [convert]::ToInt32($hex, 16)
}

function Format-EtlFile([string]$fileName, [bool]$usepktmon) {
  Write-Verbose "Format-EtlFile:$fileName"
  $error.clear()
  $outputFileName = "$fileName.csv"
  $result = $false

  if ($usepktmon) {
    Write-Host "pktmon etl2txt $fileName -o $outputFileName -m -v 5" -ForegroundColor Cyan
    $result = pktmon etl2txt $fileName -o $outputFileName -m -v 5
    $eventEntries = Read-PktmonCsvFile -fileName $outputFileName
  }
  else {
    Write-Host "netsh trace convert input=$fileName output=$outputFileName" -ForegroundColor Cyan
    $result = netsh trace convert input=$fileName output=$outputFileName
    $eventEntries = Read-NetshCsvFile -fileName $outputFileName
  }

  Write-Host "result:$result" -ForegroundColor Green

  if ($errror -or $LASTEXITCODE -ne 0) {
    Write-Host "Failed to parse ETL file: $fileName" -ForegroundColor Red
    return $null
  }

  Remove-Item -Path $outputFileName -Force

  Write-Host "Format-EtlFile: $fileName - $($eventEntries.Count) entries"
  return $eventEntries
}

function Get-Level([int]$intLevel) {
  $level = switch ($intLevel) {
    1 { return "Critical" }
    2 { return "Error" }
    3 { return "Warning" }
    4 { return "Information" }
    5 { return "Verbose" }
    default { return "Unknown" }
  }
  return $level
}

function New-Event() {
  $eventRecord = [ordered]@{
    time     = [datetime]::MinValue
    pid      = ''
    tid      = ''
    level    = ''
    keyword  = ''
    provider = ''
    event    = ''
    info     = ''
  }
  return $eventRecord
}

function Read-NetshCsvFile([string]$fileName) {
  $pattern = '[^{]+(?<json>{.+})$'

  try {
    $csvEntries = [System.Collections.ArrayList]::New()
    $streamReader = [IO.StreamReader]::New($fileName, [Text.Encoding]::Unicode)
    
    while ($streamReader.Peek() -ge 0) {
      $line = $streamReader.ReadLine()
      if (!$line) { continue }
      if ($line -match 'MSNT_SystemTrace') {
        continue
      }

      if ($line -match $pattern) {
        $jsonObject = ConvertFrom-Json -InputObject $matches['json']
        $jsonMeta = $jsonObject.meta

        $entry = New-Event
        $entry.time = [datetime]::Parse($jsonMeta.time)
        $entry.pid = $jsonMeta.pid
        $entry.tid = $jsonMeta.tid
        $entry.level = Get-Level -intLevel $jsonMeta.level
        $entry.keyword = $jsonMeta.keywords
        $entry.provider = $jsonMeta.provider
        $entry.event = $jsonMeta.event
        $entry.info = $jsonObject.Info
      }
      $csvEntries += $entry
    }

    Write-Host "Read-PktmonCsvFile: $fileName - $($csvEntries.Count) entries"
    $streamReader.Close()
    return $csvEntries
  }
  catch {
    Write-Host "exception::$($psitem.Exception.Message)`r`n$($psitem.scriptStackTrace)" -ForegroundColor Red
    $streamReader.Close()
    return $null
  }
}

function Read-PktmonCsvFile([string]$fileName) {
  $pattern = '\[(?<cpu>\d+)\](?<pid>[A-Fa-f0-9]+)\.(?<tid>[A-Fa-f0-9]+)::(?<time>.+) \[(?<provider>\w+)\] Level: (?<level>\d), Keyword: (?<keyword>.+), Event: (?<event>\w+), Info: (?<info>.+)'

  try {
    $csvEntries = [System.Collections.ArrayList]::New()
    $streamReader = [IO.StreamReader]::New($fileName, [Text.Encoding]::Unicode)
    
    while ($streamReader.Peek() -ge 0) {
      $line = $streamReader.ReadLine()
      if (!$line) { continue }
      if ($line -match 'MSNT_SystemTrace') {
        # todo: parse the header?
        continue
      }

      if ($line -match $pattern) {
        $entry = New-Event
        $entry.time = [datetime]::Parse($matches['time']).ToUniversalTime()
        $entry.pid = Convert-ToDecimal -hex $matches['pid']
        $entry.tid = Convert-ToDecimal -hex $matches['tid']
        $entry.level = Get-Level -intLevel $matches['level']
        $entry.keyword = $matches['keyword']
        $entry.provider = $matches['provider']
        $entry.event = $matches['event']
        $entry.info = $matches['info']
      }
      $csvEntries += $entry
    }

    Write-Host "Read-PktmonCsvFile: $fileName - $($csvEntries.Count) entries"
    $streamReader.Close()
    return $csvEntries
  }
  catch {
    Write-Host "exception::$($psitem.Exception.Message)`r`n$($psitem.scriptStackTrace)" -ForegroundColor Red
    $streamReader.Close()
    return $null
  }
}

function Save-ToFile([string]$outputFile) {
  if (!$outputFile) { return }

  Write-Host "Saving entries to $outputFile" -ForegroundColor Cyan

  if ((Test-Path -PathType Leaf -Path $outputFile)) {
    Write-Warning "Remove-Item -Path $outputFile -Force"
    Remove-Item -Path $outputFile -Force
  }

  if ($outputFile.ToLower().EndsWith(".json")) {
    $global:windowsUpdateEntries | ConvertTo-Json | Out-File -FilePath $outputFile
  }
  else {
    # export-csv does not format datetime correctly
    $streamWriter = [IO.StreamWriter]::New($outputFile, $false)
    $streamWriter.WriteLine("Time,PID,TID,Level,Keyword,Provider,Event,Info")

    foreach ($entry in $global:windowsUpdateEntries) {
      $line = "$($entry.time.ToString('o')),$($entry.pid),$($entry.tid),$($entry.level),$($entry.keyword),$($entry.provider),$($entry.event),$($entry.info)"
      $streamWriter.WriteLine($line)
    }
    $streamWriter.Close()
  }
}

Main