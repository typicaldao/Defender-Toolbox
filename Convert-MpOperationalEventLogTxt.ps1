# Plan: Add event ID as filter in parameter.

function Convert-MpOperationalEventLogTxt {
    param(
        [string]$Path = "$PWD\MpOperationalEvents.txt",
        [string]$OutFile = "$PWD\MpOperationalEvents.csv",
        [switch]$asCsv
    )
    
    if ((Test-Path $Path) -eq $false ) {
        Write-Host "$Path is not found! Exit"
        return
    }
    
    #Read Data
    $EventSeparator = "*****"
    $LogData = (Get-Content -Path $Path) + $EventSeparator
    
    # Some variables
    $Result = [collections.arraylist]::new()
    $i = 0 # Counter for Write-Progress
    $TotalLines = $Logdata.Count
    $Header = @{}
    $EventMessage = [text.stringbuilder]::new()

    foreach ($Line in $LogData) {
        $i++
        Write-Progress -Activity "Parsing" -Status "$i of $TotalLines lines parsed" -PercentComplete ($i / ($TotalLines) * 100)

        if ($Line.Length -lt 1) {
            continue
        }
        if ($Line.StartsWith($EventSeparator)) {
            [void]$Result.Add((Format-EventMessage $EventMessage $Header))
            [void]$EventMessage.Clear()
            continue
        }

        [void]$EventMessage.AppendLine($Line)
    }

    [void]$Result.Insert(0, (Format-Headers $Header))
    $JsonFile = convertto-json $Result -Depth 10
    if ($asCsv) {
        $Result | Export-Csv -Encoding utf8BOM -NoTypeInformation -Path $OutFile
    }
    else {
        $JsonFile | Out-File -Encoding utf8BOM -FilePath $OutFile
    }
}

function Format-EventMessage {
    param(
        [text.stringbuilder]$EventMessage,
        [hashtable]$FieldList
    )
    $FirstLine = $EventMessage.ToString().Split("`n")[0]
    $FirstLinePattern = "(?<Timestamp>.+?)`t(?<Provider>.+?)`t(?<Level>.+?)`t`t`t(?<Id>.+?)`t(?<Machine>.+?)`r"
    $FirstLineMatch = [regex]::Match($FirstLine, $FirstLinePattern)
    if ($FirstLineMatch.Success -eq $false) {
        write-warning "bad first line: $FirstLine"
        $Timestamp = "unknown"
        $Level = "unknown"
        $Id = "unknown"
        $Machine = "unknown"
    }
    else {
        $Timestamp = $FirstLineMatch.Groups["Timestamp"].Value
        $Level = $FirstLineMatch.Groups["Level"].Value
        $Id = $FirstLineMatch.Groups["Id"].Value
        $Machine = $FirstLineMatch.Groups["Machine"].Value
        # convert timestamp from g to datetime suitable for kusto ingestion
        # 6/9/2024 12:53:54 PM
        $Timestamp = [datetime]::Parse($Timestamp).ToUniversalTime().tostring("yyyy-MM-ddTHH:mm:ss.fffZ")
    }

    $EventMessage = $EventMessage.Remove(0, $FirstLine.Length + 1)
    $tabArray = $EventMessage.ToString().Split("`t")
    if ($tabArray.Count -lt 1) {
        write-warning "bad event: $EventMessage"
        return
    }

    $Description = $tabArray[0]
    $message = [ordered]@{
        EventTimestamp   = $Timestamp
        ErrorLevel       = $Level
        EventId          = $Id
        Machine          = $Machine
        EventDescription = $Description
    }

    for ($i = 1; $i -lt $tabArray.Count; $i++) {
        $LineDetails = $tabArray[$i].split(":", 2)
        if ($LineDetails.Count -lt 1) {
            write-warning "bad details: $tabArray[$i]"
            continue
        }
        $DetailsName = $LineDetails[0].Trim().replace(" ", "-")
        $DetailsValue = ""
        if ($LineDetails.Count -gt 1) {
            $DetailsValue = $LineDetails[1].Trim()
        }
        if ($message.Contains($DetailsName)) {
            Write-Warning "Duplicate key: $DetailsName : $DetailsValue"
        }
        else {
            [void]$message.Add($DetailsName, $DetailsValue)
        }
        if ($FieldList.Contains($DetailsName) -eq $false) {
            [void]$FieldList.Add($DetailsName, $DetailsName)
        }
    }
    return $message
}

function Format-Headers {
    param(
        [hashtable]$HeaderDictionary
    )

    $Header = [ordered]@{
        EventTimestamp   = "Timestamp"
        ErrorLevel       = "Level"
        EventId          = "Id"
        Machine          = "Machine"
        EventDescription = "Description"
    }

    foreach ($kvp in $HeaderDictionary.GetEnumerator() | Sort-Object Name) {
        if ($Header.Contains($kvp.name)) {
            Write-Warning "Duplicate key: $kvp.name : $kvp.value"
        }
        else {
            [void]$Header.Add($kvp.name, $kvp.value)
        }
    }
    return $Header
}