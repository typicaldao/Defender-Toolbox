# Plan: Add event ID as filter in parameter.

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
        $result | Export-Csv -Encoding utf8BOM -NoHeader -NoTypeInformation -Path $outFile
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
            $message.EventDetails = $eventDetailsJson | ConvertTo-Json -Compress
        }
        else {
            $message.EventDetails = ([string]::join("", $eventMessage.ToArray()).Replace("`r`n", "`t"))
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
}