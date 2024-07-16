function Convert-MpOperationalEventLogTxt {
    param(
        [string]$Path="$PWD\MpOperationalEvents.txt",
        [string]$OutFile="$PWD\MpOperationalEvents.csv"
    )
    
    if ((Test-Path $Path) -eq $false ) {
        Write-Host "MpOperationalEvents.txt is not found! Exiting."
        return
    }
    
    #Read Data
    $LogData = (Get-Content -Path $Path) + "*****"
    
    # Some variables
    $Counter = -1
    $Details = ""
    $Result = [array][PSCustomObject]@() 
    $Description = ""
    
    $i = 0 # Counter for Write-Progress
    $TotalLines = $Logdata.Count
    
    $LogData | ForEach-Object {
        if ($_ -match '^[*]+') {
            if ($Counter -ne -1) {
               $Result += [PSCustomObject]@{
                Time = $Timestamp
                ErrorLevel = $Level
                EventId = $Id
                EventDescription = $Description
                EventDetails = $Details
               }
            }
            # Initialization of event block
            $Details = ""
            $Counter = 0 
        } 
        # Parse the first line of event info
        if ($Counter -eq 1) {  
            $Timestamp = $_.Substring(0,22) 
            $Level,$Id = $_.Split()[5],$_.Split()[8] 
        } 
        # One-line description
        elseif ($Counter -eq 2) {
            $Description = $_
        }
        # Multi-lines details
        elseif ($Counter -gt 2) {
            $Details += $_ + "`n"
        }
        $Counter++ 
        $i++
        Write-Progress -Activity "Parsing" -Status "$i of $TotalLines lines parsed" -PercentComplete ($i/($TotalLines)*100)
    }
    
    $Result | Export-Csv -Path $OutFile
}
Export-ModuleMember -Function Convert-MpOperationalEventLogTxt

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